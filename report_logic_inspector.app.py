"""Streamlit app to inspect Excel/CSV reports for embedded logic (macros, queries, connections)."""

from __future__ import annotations

import io
import re
import textwrap
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

try:  # Optional dependency for extracting VBA code
    from oletools.olevba import VBA_Parser  # type: ignore
except Exception:  # pragma: no cover - graceful degradation when oletools missing
    VBA_Parser = None  # type: ignore


st.set_page_config(
    page_title="Report Logic Inspector",
    page_icon="🧭",
    layout="wide",
)


@dataclass
class MacroModule:
    name: str
    path: str
    lines: int
    summary: List[str]
    code: str


@dataclass
class MacroSummary:
    found: bool = False
    modules: List[MacroModule] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    requires_dependency: bool = False


@dataclass
class PowerQuerySummary:
    name: str
    load_enabled: Optional[bool]
    summary: List[str]
    script: str


@dataclass
class ConnectionSummary:
    name: str
    connection_type: str
    summary: List[str]
    properties: Dict[str, str]


@dataclass
class ReportAnalysis:
    name: str
    file_type: str
    size_bytes: int
    analysis_ts: datetime
    metadata: Dict[str, str]
    sheets: Sequence[str]
    macros: MacroSummary
    power_queries: List[PowerQuerySummary]
    connections: List[ConnectionSummary]
    power_pivot_present: bool
    extra_notes: List[str]
    csv_preview: Optional[pd.DataFrame] = None


POWER_QUERY_NS = {"pq": "http://schemas.microsoft.com/office/PowerQuery/2013/Main"}


def human_size(num: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if abs(num) < 1024.0:
            return f"{num:.1f} {unit}" if unit != "B" else f"{num} {unit}"
        num /= 1024.0
    return f"{num:.2f} TB"


def summarize_vba_code(code: str) -> List[str]:
    lowered = code.lower()
    insights: List[str] = []
    if "workbooks.open" in lowered or "filedialog" in lowered:
        insights.append("Opens external workbooks or files.")
    if "querytables.add" in lowered or "listobjects" in lowered:
        insights.append("Refreshes or injects data connections/tables.")
    if "ado" in lowered or "connection" in lowered:
        insights.append("Interacts with databases via ADO/connection objects.")
    if "sheet" in lowered or "cells" in lowered:
        insights.append("Manipulates worksheet cells or ranges.")
    if "auto_open" in lowered or "workbook_open" in lowered:
        insights.append("Runs automatically when the workbook opens.")
    if not insights:
        insights.append("General VBA logic detected; review code block for specifics.")
    return insights


def extract_macros(workbook_name: str, workbook_bytes: bytes) -> MacroSummary:
    summary = MacroSummary()
    if VBA_Parser is None:
        # Quick check if a macro project exists before notifying about the dependency.
        if contains_vba_project(workbook_bytes):
            summary.found = True
            summary.requires_dependency = True
            summary.errors.append(
                "vbaProject.bin detected. Install 'oletools' to extract module code (pip install oletools)."
            )
        return summary

    try:
        parser = VBA_Parser(workbook_name, data=workbook_bytes)
    except Exception as exc:  # pragma: no cover - defensive
        summary.errors.append(f"VBA parser error: {exc}")
        return summary

    if not parser.detect_vba_macros():
        return summary

    summary.found = True
    try:
        for (_, stream_path, module_name, vba_code) in parser.extract_macros():
            code_text = vba_code.decode("utf-8", errors="ignore") if isinstance(vba_code, bytes) else vba_code
            summary.modules.append(
                MacroModule(
                    name=module_name or stream_path.split("/")[-1],
                    path=stream_path,
                    lines=len(code_text.splitlines()),
                    summary=summarize_vba_code(code_text),
                    code=code_text,
                )
            )
    except Exception as exc:  # pragma: no cover
        summary.errors.append(f"Failed to extract macros: {exc}")
    finally:
        parser.close()
    return summary


def contains_vba_project(workbook_bytes: bytes) -> bool:
    with io.BytesIO(workbook_bytes) as buffer:
        if zipfile.is_zipfile(buffer):
            with zipfile.ZipFile(buffer) as zf:
                return any(name.lower().endswith("vbaProject.bin".lower()) for name in zf.namelist())
    # Legacy XLS or binary workbook signature (D0 CF 11 E0) indicates possible VBA.
    return workbook_bytes[:4] == b"\xd0\xcf\x11\xe0"


def read_xml(zf: zipfile.ZipFile, path: str) -> Optional[ET.Element]:
    try:
        with zf.open(path) as handle:
            data = handle.read()
        return ET.fromstring(data)
    except KeyError:
        return None
    except ET.ParseError:
        return None


def extract_metadata(zf: zipfile.ZipFile) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    core = read_xml(zf, "docProps/core.xml")
    if core is not None:
        for elem in core:
            tag = elem.tag.split("}")[-1]
            if elem.text:
                meta[tag.capitalize()] = elem.text
    app = read_xml(zf, "docProps/app.xml")
    if app is not None:
        security = app.find("{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Security")
        if security is not None and security.text:
            meta["Security"] = security.text
        company = app.find("{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company")
        if company is not None and company.text:
            meta["Company"] = company.text
    return meta


def extract_sheet_names(zf: zipfile.ZipFile) -> List[str]:
    names: List[str] = []
    workbook = read_xml(zf, "xl/workbook.xml")
    if workbook is None:
        return names
    for sheet in workbook.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
        name = sheet.get("name")
        if name:
            names.append(name)
    return names


def summarize_m_script(script: str) -> List[str]:
    text = script.lower()
    insights: List[str] = []

    def extract_literal(pattern: str) -> Optional[str]:
        match = re.search(pattern, script, re.IGNORECASE)
        if match:
            return match.group(1).strip().strip("\"")
        return None

    if "sql.database" in text:
        server = extract_literal(r"Sql\.Database\(([^,]+)")
        database = extract_literal(r"Sql\.Database\([^,]+,([^\)]+)\)")
        details = "SQL Server source"
        if server:
            details += f" on {server}"
        if database:
            details += f" using database {database}"
        insights.append(details + ".")
    if "sqldatabase" in text and "" not in insights:
        pass  # handled above
    if "odbc.query" in text or "ole db" in text:
        insights.append("Runs custom SQL via ODBC/OLE DB connection.")
    if "sharepoint" in text:
        insights.append("Connects to SharePoint data (files or lists).")
    if "web.contents" in text:
        insights.append("Calls a web API using Web.Contents.")
    if "csv.document" in text or "excel.workbook" in text:
        insights.append("Reads structured data from a local file (CSV/Excel).")
    if "table.combine" in text:
        insights.append("Combines multiple tables/queries into one result.")
    if "merge" in text or "joinkind" in text:
        insights.append("Performs join/merge operations between tables.")
    if "group" in text and "table.group" in text:
        insights.append("Aggregates data using Table.Group.")
    if "table.transformcolumn" in text:
        insights.append("Applies column type or value transformations.")
    if not insights:
        insights.append("Power Query steps detected; review script for detailed logic.")
    return insights


def extract_power_queries(zf: zipfile.ZipFile) -> List[PowerQuerySummary]:
    summaries: List[PowerQuerySummary] = []
    for name in zf.namelist():
        if not name.startswith("xl/queries/") or not name.lower().endswith(".xml"):
            continue
        try:
            with zf.open(name) as handle:
                xml_text = handle.read()
        except KeyError:
            continue
        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError:
            continue

        ns_uri = POWER_QUERY_NS["pq"]
        if root.tag.startswith("{"):
            ns_uri = root.tag.split("}")[0][1:]
        ns = {"pq": ns_uri}

        query_name = root.findtext("pq:Name", namespaces=ns) or Path(name).stem
        script = root.findtext("pq:Formula", namespaces=ns) or ""
        load = root.findtext("pq:LoadEnabled", namespaces=ns)
        load_enabled: Optional[bool] = None
        if load is not None:
            load_enabled = load.lower() == "true"
        summaries.append(
            PowerQuerySummary(
                name=query_name,
                load_enabled=load_enabled,
                summary=summarize_m_script(script),
                script=script.strip(),
            )
        )
    return summaries


def interpret_connection_string(conn_str: str) -> List[str]:
    insights: List[str] = []
    server_match = re.search(r"data source=([^;]+)", conn_str, flags=re.IGNORECASE)
    if server_match:
        insights.append(f"Data source/server: {server_match.group(1)}")
    db_match = re.search(r"initial catalog=([^;]+)", conn_str, flags=re.IGNORECASE)
    if db_match:
        insights.append(f"Initial catalog/database: {db_match.group(1)}")
    provider_match = re.search(r"provider=([^;]+)", conn_str, flags=re.IGNORECASE)
    if provider_match:
        insights.append(f"Provider: {provider_match.group(1)}")
    lower_conn = conn_str.lower()
    if "oledb" in lower_conn:
        insights.append("Uses an OLE DB provider.")
    if "odbc;" in lower_conn:
        insights.append("Uses an ODBC DSN or driver.")
    if "password" in lower_conn:
        insights.append("Contains embedded credentials – review security.")
    if not insights:
        insights.append("Connection string provided; inspect for server/database details.")
    return insights


def extract_connections(zf: zipfile.ZipFile) -> List[ConnectionSummary]:
    node = read_xml(zf, "xl/connections.xml")
    if node is None:
        return []
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    summaries: List[ConnectionSummary] = []
    for conn in node.findall(f"{ns}connection"):
        name = conn.get("name") or conn.get("id") or "Unnamed connection"
        conn_type = conn.get("type") or conn.get("refreshOnLoad") or "Unknown"
        props: Dict[str, str] = {}
        inner = conn.find(f"{ns}dbPr") or conn.find(f"{ns}odcPr") or conn.find(f"{ns}webPr")
        summary_parts: List[str] = []
        if inner is not None:
            for key, value in inner.attrib.items():
                props[key] = value
            conn_str = inner.get("connection") or inner.get("dbCommand") or ""
            if conn_str:
                summary_parts.extend(interpret_connection_string(conn_str))
            if inner.tag.endswith("dbPr") and inner.get("command"):
                summary_parts.append(f"Runs command: {inner.get('command')[:80]}...")
        if not summary_parts:
            summary_parts.append("Connection metadata present; inspect attributes for detail.")
        summaries.append(
            ConnectionSummary(
                name=name,
                connection_type=conn_type,
                summary=summary_parts,
                properties=props,
            )
        )
    return summaries


def detect_power_pivot(zf: zipfile.ZipFile) -> bool:
    return any(name.startswith("xl/model") or name.startswith("xl/powerPivot") for name in zf.namelist())


def analyse_excel_file(name: str, data: bytes) -> ReportAnalysis:
    with io.BytesIO(data) as buffer:
        if not zipfile.is_zipfile(buffer):
            extra = [
                "Workbook is not in the OpenXML (.xlsx/.xlsm) format; legacy .xls parsing is limited.",
                "Consider converting to .xlsm/.xlsx to expose internal XML metadata.",
            ]
            return ReportAnalysis(
                name=name,
                file_type="Legacy workbook",
                size_bytes=len(data),
                analysis_ts=datetime.utcnow(),
                metadata={},
                sheets=[],
                macros=extract_macros(name, data),
                power_queries=[],
                connections=[],
                power_pivot_present=False,
                extra_notes=extra,
            )
        buffer.seek(0)
        with zipfile.ZipFile(buffer) as zf:
            return ReportAnalysis(
                name=name,
                file_type="Excel OpenXML workbook",
                size_bytes=len(data),
                analysis_ts=datetime.utcnow(),
                metadata=extract_metadata(zf),
                sheets=extract_sheet_names(zf),
                macros=extract_macros(name, data),
                power_queries=extract_power_queries(zf),
                connections=extract_connections(zf),
                power_pivot_present=detect_power_pivot(zf),
                extra_notes=[],
            )


def analyse_csv_file(name: str, data: bytes) -> ReportAnalysis:
    preview_df: Optional[pd.DataFrame] = None
    notes: List[str] = []
    try:
        preview_df = pd.read_csv(io.BytesIO(data), nrows=20)
        notes.append("CSV preview limited to first 20 rows for context.")
    except Exception as exc:
        notes.append(f"Unable to read CSV preview: {exc}")
    return ReportAnalysis(
        name=name,
        file_type="Delimited text",
        size_bytes=len(data),
        analysis_ts=datetime.utcnow(),
        metadata={},
        sheets=[],
        macros=MacroSummary(),
        power_queries=[],
        connections=[],
        power_pivot_present=False,
        extra_notes=notes,
        csv_preview=preview_df,
    )


@st.cache_data(show_spinner=False)
def analyse_uploaded_file(name: str, data: bytes) -> ReportAnalysis:
    suffix = Path(name).suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltm", ".xlsb", ".xls"}:
        return analyse_excel_file(name, data)
    if suffix in {".csv", ".txt", ".tsv"}:
        return analyse_csv_file(name, data)
    if zipfile.is_zipfile(io.BytesIO(data)):
        return analyse_excel_file(name, data)
    return ReportAnalysis(
        name=name,
        file_type="Unsupported/unknown",
        size_bytes=len(data),
        analysis_ts=datetime.utcnow(),
        metadata={},
        sheets=[],
        macros=MacroSummary(),
        power_queries=[],
        connections=[],
        power_pivot_present=False,
        extra_notes=["File type not recognized for automated inspection."],
    )


def render_macro_section(summary: MacroSummary) -> None:
    if not summary.found:
        st.info("No VBA macro project detected.")
        if summary.errors:
            for err in summary.errors:
                st.caption(err)
        return

    st.success(f"Macro project detected with {len(summary.modules) or 'unknown'} module(s).")
    if summary.requires_dependency:
        for err in summary.errors:
            st.warning(err)
        return
    if summary.errors:
        for err in summary.errors:
            st.warning(err)

    for module in summary.modules:
        with st.expander(f"Module: {module.name} ({module.lines} lines)", expanded=False):
            if module.summary:
                st.markdown("**Translation**")
                for insight in module.summary:
                    st.write(f"- {insight}")
            st.markdown("**Code Preview**")
            st.code(module.code, language="vb")


def render_query_section(queries: Sequence[PowerQuerySummary]) -> None:
    if not queries:
        st.info("No Power Query definitions found.")
        return
    for query in queries:
        title = query.name
        if query.load_enabled is not None:
            load_status = "Loads to sheet/model" if query.load_enabled else "Connection only"
            title += f" · {load_status}"
        with st.expander(title, expanded=False):
            st.markdown("**Translation**")
            for insight in query.summary:
                st.write(f"- {insight}")
            if query.script:
                st.markdown("**M Script**")
                st.code(query.script, language="m")


def render_connection_section(connections: Sequence[ConnectionSummary]) -> None:
    if not connections:
        st.info("No workbook connection metadata detected.")
        return
    for conn in connections:
        with st.expander(f"Connection: {conn.name}", expanded=False):
            st.markdown(f"**Type:** {conn.connection_type}")
            if conn.summary:
                st.markdown("**Translation**")
                for insight in conn.summary:
                    st.write(f"- {insight}")
            if conn.properties:
                st.markdown("**Raw Properties**")
                st.json(conn.properties)


def render_analysis(result: ReportAnalysis) -> None:
    st.markdown(f"### {result.name}")
    cols = st.columns(4)
    cols[0].metric("File type", result.file_type)
    cols[1].metric("Size", human_size(result.size_bytes))
    cols[2].metric("Analysed", result.analysis_ts.strftime("%Y-%m-%d %H:%M:%S UTC"))
    cols[3].metric("Sheets", len(result.sheets))

    if result.metadata:
        with st.expander("Workbook metadata", expanded=False):
            st.json(result.metadata)
    if result.sheets:
        st.caption("Worksheets detected: " + ", ".join(result.sheets))

    st.markdown("#### Macros / VBA")
    render_macro_section(result.macros)

    st.markdown("#### Power Query (Get & Transform)")
    render_query_section(result.power_queries)

    st.markdown("#### Data Connections")
    render_connection_section(result.connections)

    if result.power_pivot_present:
        st.warning("Power Pivot / data model artifacts detected – inspect with Power Pivot add-in.")

    if result.csv_preview is not None:
        st.markdown("#### CSV Preview")
        st.dataframe(result.csv_preview, use_container_width=True)

    if result.extra_notes:
        st.markdown("#### Notes")
        for note in result.extra_notes:
            st.write(f"- {note}")


def main() -> None:
    st.title("Report Logic Inspector")
    st.caption(
        "Upload Excel or CSV reports to discover embedded macros, Power Queries, data connections, and other logic."
    )

    with st.sidebar:
        st.markdown("### How it works")
        st.write(
            textwrap.dedent(
                """
                • Looks inside OpenXML Excel workbooks for VBA, Power Query, and connection metadata.
                • Summarizes detected logic so you can understand data lineage and automation steps.
                • Optional: install `oletools` to extract full VBA source code for review.
                """
            )
        )

    uploaded_files = st.file_uploader(
        "Select one or more report files",
        type=["xlsx", "xlsm", "xls", "xltm", "xlsb", "csv", "txt", "tsv"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Upload at least one Excel or CSV report to begin analysis.")
        return

    for uploaded in uploaded_files:
        file_bytes = uploaded.getvalue()
        with st.spinner(f"Analysing {uploaded.name}…"):
            result = analyse_uploaded_file(uploaded.name, file_bytes)
        render_analysis(result)
        st.markdown("---")


if __name__ == "__main__":
    main()
