import json
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple, Literal, Set

import streamlit as st


# Configure the Streamlit page before any UI elements render.
st.set_page_config(page_title="EXP Analysis Tool", layout="wide")


@dataclass
class ComparisonRecord:
    # Lightweight wrapper that normalizes key fields from a comparisons.jsonl row.
    raw: dict
    line_number: int

    @property
    def stored_procedure(self) -> Optional[str]:
        keys = (
            "stored_procedure_name",
            "storedProcedureName",
            "stored_proc_name",
            "procedure_name",
        )
        for key in keys:
            value = self.raw.get(key)
            if isinstance(value, str) and value.strip():
                return value
        return None

    @property
    def execution_hash(self) -> Optional[str]:
        keys = ("execution_hash", "executionHash", "execution_id")
        for key in keys:
            value = self.raw.get(key)
            if isinstance(value, str) and value.strip():
                return value
        return None

    @property
    def join_columns(self) -> Optional[Sequence[str]]:
        value = self.raw.get("join_columns")
        if value is None:
            return None
        if isinstance(value, list):
            return value
        if isinstance(value, str):
            cleaned = [col.strip() for col in value.split(",") if col.strip()]
            return cleaned or None
        return None

    @property
    def join_columns_raw(self):
        return self.raw.get("join_columns")

    def is_join_columns_missing(self) -> bool:
        raw = self.join_columns_raw
        if raw is None:
            return True
        if isinstance(raw, str):
            return raw.strip() in ("", "0")
        if isinstance(raw, list):
            return len([col for col in raw if str(col).strip()]) == 0
        return True


@dataclass
class AnalysisIssue:
    line: int
    severity: Literal["info", "warning", "error"]
    message: str
    recommendation: Optional[str] = None
    context: Optional[str] = None


def resolve_path(base: Path, value: str) -> Path:
    # Resolve user-provided paths relative to the working directory/panel entries.
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base / path
    return path


def load_jsonl(path: Path) -> Tuple[List[ComparisonRecord], List[str]]:
    records: List[ComparisonRecord] = []
    errors: List[str] = []
    if not path.exists():
        return records, [f"comparisons file not found at {path}"]
    with path.open("r", encoding="utf-8") as handle:
        for idx, line in enumerate(handle, start=1):
            content = line.strip()
            if not content:
                continue
            try:
                payload = json.loads(content)
            except json.JSONDecodeError as exc:
                errors.append(f"Line {idx}: {exc}")
                continue
            records.append(ComparisonRecord(raw=payload, line_number=idx))
    return records, errors


def write_jsonl(path: Path, records: Iterable[ComparisonRecord]) -> None:
    # Persist comparison records back to disk in JSONL format.
    with path.open("w", encoding="utf-8") as handle:
        for record in records:
            handle.write(json.dumps(record.raw))
            handle.write("\n")


def read_text(path: Path) -> str:
    # Load helper text files (filters) while tolerating missing paths.
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def write_text(path: Path, content: str) -> None:
    # Save edited filter files, creating directories as needed.
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_bytes(path: Path, content: bytes) -> None:
    # Persist uploaded EXP artifacts to disk within the working tree.
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)


def analyze_tsql(content: str) -> List[AnalysisIssue]:
    # Heuristic scan of a T-SQL script to highlight Snowflake incompatibilities.
    issues: List[AnalysisIssue] = []
    lines = content.splitlines()

    def add_issue(
        line_index: int,
        severity: Literal["info", "warning", "error"],
        message: str,
        recommendation: Optional[str] = None,
        context: Optional[str] = None,
    ) -> None:
        issues.append(
            AnalysisIssue(
                line=line_index,
                severity=severity,
                message=message,
                recommendation=recommendation,
                context=context,
            )
        )

    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()
        upper = stripped.upper()
        context = stripped or None

        if not stripped:
            continue

        if upper == "GO":
            add_issue(
                idx,
                "warning",
                "Batch separator `GO` is not supported in Snowflake.",
                "Remove `GO` statements and execute statements sequentially.",
                context,
            )

        if re.match(r"^\s*USE\s+", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "`USE` statements are not supported in Snowflake sessions.",
                "Establish database/schema with `USE DATABASE` / `USE SCHEMA` outside the script or rely on the session context.",
                context,
            )

        if re.search(r"\bIDENTITY\s*\(", line, re.IGNORECASE):
            add_issue(
                idx,
                "error",
                "IDENTITY columns are not supported in Snowflake.",
                "Replace with a sequence and default value (e.g., `sequence_name.NEXTVAL`).",
                context,
            )

        if re.search(r"\bTOP\s+\d+", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "`TOP` syntax is not available in Snowflake.",
                "Rewrite the query to use `LIMIT n` (optionally with `ORDER BY`).",
                context,
            )

        if re.search(r"\[[^\]]+\]", line):
            add_issue(
                idx,
                "info",
                "Bracket-delimited identifiers should be quoted with double quotes in Snowflake.",
                'Replace `[identifier]` with `"identifier"`.',
                context,
            )

        if re.search(r"\bNVARCHAR\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "NVARCHAR is not a Snowflake data type.",
                "Use `VARCHAR` or `VARCHAR(n)`.",
                context,
            )

        if re.search(r"\bNCHAR\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "NCHAR is not a Snowflake data type.",
                "Use `CHAR` or `CHAR(n)`.",
                context,
            )

        if re.search(r"\bDATETIME\b", line, re.IGNORECASE) or re.search(r"\bSMALLDATETIME\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "DATETIME/SMALLDATETIME types should be mapped to Snowflake TIMESTAMP variants.",
                "Use `TIMESTAMP_NTZ`, `TIMESTAMP_LTZ`, or `TIMESTAMP_TZ` as appropriate.",
                context,
            )

        if re.search(r"\bMONEY\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "MONEY type is not available in Snowflake.",
                "Use `NUMBER(38, 4)` or a numeric type that preserves required scale.",
                context,
            )

        if re.search(r"\bBIT\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "BIT type is not supported in Snowflake.",
                "Use `BOOLEAN` instead.",
                context,
            )

        if re.search(r"\bUNIQUEIDENTIFIER\b", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "UNIQUEIDENTIFIER has no direct Snowflake equivalent.",
                "Represent GUID values as `VARCHAR` or use Snowflake UUID functions.",
                context,
            )

        if re.search(r"\bGETDATE\s*\(\s*\)", line, re.IGNORECASE) or re.search(r"\bSYSDATETIME\s*\(\s*\)", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "Date/time functions differ between SQL Server and Snowflake.",
                "Use `CURRENT_TIMESTAMP` or `CURRENT_DATE` as needed.",
                context,
            )

        if re.search(r"\bISNULL\s*\(", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "ISNULL behaves differently in Snowflake.",
                "Replace with `COALESCE` to support multiple arguments.",
                context,
            )

        if re.search(r"\bCONVERT\s*\(", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "CONVERT function is not available in Snowflake.",
                "Use `CAST` or `TO_<TYPE>` helpers.",
                context,
            )

        if re.search(r"\bLEN\s*\(", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "LEN should be replaced in Snowflake.",
                "Use `LENGTH` or `CHAR_LENGTH`.",
                context,
            )

        if re.search(r"@@IDENTITY|SCOPE_IDENTITY", line, re.IGNORECASE):
            add_issue(
                idx,
                "warning",
                "SQL Server identity retrieval functions are unavailable in Snowflake.",
                "Use sequences and capture values explicitly (e.g., `LAST_VALUE` from sequence).",
                context,
            )

        if re.search(r"\bEXEC\s+", line, re.IGNORECASE):
            add_issue(
                idx,
                "info",
                "Stored procedure execution uses different syntax in Snowflake.",
                "Replace with `CALL procedure_name(arguments);`.",
                context,
            )

        if re.search(r"#\w+", line):
            add_issue(
                idx,
                "warning",
                "Temporary tables defined with `#name` are SQL Server specific.",
                "Use Snowflake temporary tables (`CREATE TEMP TABLE`) without the hash prefix.",
                context,
            )

    return issues


def transform_tsql_to_snow(content: str) -> Tuple[str, List[str]]:
    # Apply conservative text transformations to produce a SnowSQL starting point.
    notes: List[str] = []
    seen: Set[str] = set()

    def add_note(text: str) -> None:
        if text not in seen:
            seen.add(text)
            notes.append(text)

    lines: List[str] = []
    for line in content.splitlines():
        if line.strip().upper() == "GO":
            add_note("Removed `GO` batch separators.")
            continue
        lines.append(line)

    text = "\n".join(lines)

    bracket_pattern = re.compile(r"\[([^\]]+)\]")
    if bracket_pattern.search(text):
        add_note("Converted bracket-delimited identifiers to double quotes.")
        text = bracket_pattern.sub(lambda match: f'"{match.group(1)}"', text)

    def apply(pattern: str, replacement: str, note: str, flags: int = re.IGNORECASE) -> None:
        nonlocal text
        new_text, count = re.subn(pattern, replacement, text, flags=flags)
        if count:
            add_note(note)
            text = new_text

    apply(r"\bGETDATE\s*\(\s*\)", "CURRENT_TIMESTAMP", "Replaced GETDATE() with CURRENT_TIMESTAMP.")
    apply(r"\bSYSDATETIME\s*\(\s*\)", "CURRENT_TIMESTAMP", "Replaced SYSDATETIME() with CURRENT_TIMESTAMP.")
    apply(r"\bISNULL\s*\(", "COALESCE(", "Replaced ISNULL with COALESCE.")
    apply(r"\bNVARCHAR\s*\(", "VARCHAR(", "Replaced NVARCHAR with VARCHAR.")
    apply(r"\bNVARCHAR\b", "VARCHAR", "Replaced NVARCHAR with VARCHAR.")
    apply(r"\bNCHAR\s*\(", "CHAR(", "Replaced NCHAR with CHAR.")
    apply(r"\bNCHAR\b", "CHAR", "Replaced NCHAR with CHAR.")
    apply(r"\bDATETIME\b", "TIMESTAMP_NTZ", "Mapped DATETIME to TIMESTAMP_NTZ.")
    apply(r"\bSMALLDATETIME\b", "TIMESTAMP_NTZ", "Mapped SMALLDATETIME to TIMESTAMP_NTZ.")
    apply(r"\bMONEY\b", "NUMBER(38,4)", "Mapped MONEY to NUMBER(38,4).")
    apply(r"\bBIT\b", "BOOLEAN", "Mapped BIT to BOOLEAN.")
    apply(r"\bUNIQUEIDENTIFIER\b", "VARCHAR", "Represented UNIQUEIDENTIFIER as VARCHAR.")
    apply(r"\bLEN\s*\(", "LENGTH(", "Replaced LEN with LENGTH.")

    return text, notes


def run_datacompare(
    action: str,
    binary: Path,
    config_path: Path,
    comparison_filter: Path,
    workdir: Path,
    extra_args: Optional[Sequence[str]] = None,
) -> Tuple[int, str, str, List[str]]:
    # Centralized subprocess runner for datacompare commands.
    command = [
        str(binary),
        "--action",
        action,
        "--config",
        str(config_path),
        "--comparison-filter-file",
        str(comparison_filter),
    ]
    if extra_args:
        command.extend(extra_args)
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            cwd=str(workdir),
            check=False,
        )
    except FileNotFoundError as exc:
        return 127, "", str(exc), command
    return result.returncode, result.stdout, result.stderr, command


def detect_match_success(output: str) -> bool:
    # Look for the "Match success" signature in datacompare output.
    pattern = re.compile(r"\bMatch\s+success\b", re.IGNORECASE)
    return bool(pattern.search(output))


def detect_duplicates(output: str) -> Optional[bool]:
    # Parse duplicate reporting line to guide next steps.
    match = re.search(r"Any duplicates on match values:\s*(Yes|No)", output, re.IGNORECASE)
    if not match:
        return None
    return match.group(1).lower() == "yes"


def suggest_columns(output: str) -> List[str]:
    # Scrape likely column names from find_keys output to aid join_columns selection.
    bracket_pattern = re.compile(r"\[(.*?)\]")
    quoted_pattern = re.compile(r'"([^"]+)"')
    candidates = set()
    for match in bracket_pattern.findall(output):
        parts = [item.strip(" '\"") for item in match.split(",")]
        for part in parts:
            if is_candidate_column(part):
                candidates.add(part)
    for match in quoted_pattern.findall(output):
        if is_candidate_column(match):
            candidates.add(match)
    token_pattern = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")
    for token in token_pattern.findall(output):
        if is_candidate_column(token):
            candidates.add(token)
    return sorted(candidates)


def is_candidate_column(token: str) -> bool:
    reserved = {
        "Candidate",
        "Candidates",
        "Keys",
        "Key",
        "Match",
        "success",
        "Yes",
        "No",
        "Duplicates",
        "NULL",
    }
    stripped = token.strip().strip("'\"")
    if not stripped or stripped in reserved:
        return False
    if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", stripped):
        return False
    return True


def render_path_inputs(base_default: Path) -> Tuple[Path, Path, Path, Path, Path, Path]:
    # Sidebar controls for configuring binary/config/filter file locations.
    st.sidebar.header("Paths")
    workdir_input = st.sidebar.text_input("Working directory", value=str(base_default))
    workdir_path = Path(workdir_input).expanduser()

    binary_input = st.sidebar.text_input("datacompare binary", value="./datacompare")
    comparisons_input = st.sidebar.text_input("comparisons.jsonl", value="config/comparisons.jsonl")
    comparison_filter_input = st.sidebar.text_input("comparison_filter.txt", value="config/comparison_filter.txt")
    compare_filter_input = st.sidebar.text_input("compare_filter.txt", value="config/compare_filter.txt")
    config_input = st.sidebar.text_input(
        "config yaml",
        value="config/config_snow_dev_sql_prod.yaml",
    )

    datacompare_path = resolve_path(workdir_path, binary_input)
    comparisons_path = resolve_path(workdir_path, comparisons_input)
    comparison_filter_path = resolve_path(workdir_path, comparison_filter_input)
    compare_filter_path = resolve_path(workdir_path, compare_filter_input)
    config_path = resolve_path(workdir_path, config_input)
    return (
        workdir_path,
        datacompare_path,
        comparisons_path,
        comparison_filter_path,
        compare_filter_path,
        config_path,
    )


def render_exp_upload(workdir_path: Path) -> Tuple[Optional[Path], Optional[str]]:
    # Allow analysts to upload an EXP text file for reference during analysis.
    st.sidebar.subheader("EXP Upload")
    uploaded_file = st.sidebar.file_uploader("Upload EXP file (.txt)", type=["txt"])

    exp_path: Optional[Path] = None
    exp_content: Optional[str] = None

    if uploaded_file is not None:
        data = uploaded_file.getvalue()
        filename = uploaded_file.name or "uploaded_exp.txt"
        safe_name = Path(filename).name
        exp_path = workdir_path / "exp_uploads" / safe_name
        write_bytes(exp_path, data)
        exp_content = data.decode("utf-8", errors="replace")
        st.sidebar.success(f"Saved EXP file to {exp_path}")
        st.session_state["exp_file_path"] = str(exp_path)
        st.session_state["exp_file_content"] = exp_content
    else:
        stored_path = st.session_state.get("exp_file_path")
        stored_content = st.session_state.get("exp_file_content")
        if stored_path:
            exp_path = Path(stored_path)
        if stored_content:
            exp_content = stored_content

    if exp_path:
        st.sidebar.caption(f"Active EXP file: {exp_path}")

    return exp_path, exp_content


def pick_record(records: List[ComparisonRecord]) -> Optional[ComparisonRecord]:
    # Let the analyst choose which comparison entry to work on.
    if not records:
        st.error("No comparison records available.")
        return None
    options = list(range(len(records)))

    def format_option(idx: int) -> str:
        record = records[idx]
        procedure = record.stored_procedure or "<unknown stored procedure>"
        execution = record.execution_hash or "<no execution hash>"
        join_status = "set" if not record.is_join_columns_missing() else "missing"
        return f"{procedure} | exec hash: {execution} | join_columns: {join_status}"

    selection = st.selectbox(
        "Stored procedure",
        options,
        format_func=format_option,
        index=0,
    )
    return records[selection]


def update_join_columns(
    comparisons_path: Path,
    records: List[ComparisonRecord],
    target_record: ComparisonRecord,
    new_columns: Sequence[str],
) -> None:
    # Overwrite join_columns for the selected record while leaving others intact.
    new_values: Sequence[str] = [col.strip() for col in new_columns if col.strip()]
    for record in records:
        if record is target_record:
            original = record.join_columns_raw
            if isinstance(original, list):
                record.raw["join_columns"] = list(new_values)
            elif isinstance(original, str):
                record.raw["join_columns"] = ",".join(new_values)
            else:
                record.raw["join_columns"] = list(new_values)
            break
    write_jsonl(comparisons_path, records)


def render_filter_editor(label: str, path: Path) -> None:
    # Inline editor for maintaining comparison/compare filter text files.
    st.subheader(label)
    content = read_text(path)
    updated = st.text_area(f"{path}", value=content, height=150)
    if st.button(f"Save {label}", key=f"save_{label}"):
        write_text(path, updated)
        st.success(f"Saved {label} to {path}")


def render_command_output(stdout: str, stderr: str) -> None:
    if stdout:
        st.markdown("**stdout**")
        st.code(stdout, language="text")
    if stderr:
        st.markdown("**stderr**")
        st.code(stderr, language="text")


def main() -> None:
    # Primary Streamlit view orchestrating the EXP analysis workflow.
    base_default = Path.cwd()
    (
        workdir_path,
        datacompare_path,
        comparisons_path,
        comparison_filter_path,
        compare_filter_path,
        config_path,
    ) = render_path_inputs(base_default)
    exp_path, exp_content = render_exp_upload(workdir_path)

    st.title("EXP Analysis Workflow Assistant")
    st.caption("Guide and automate the EXP data comparison workflow.")

    if exp_content:
        st.subheader("Uploaded EXP File")
        if exp_path:
            st.caption(f"Stored at: {exp_path}")
        st.text_area("EXP contents", value=exp_content, height=200, key="exp_source_view")
        st.download_button(
            label="Download uploaded EXP",
            data=exp_content,
            file_name=(exp_path.name if exp_path else "uploaded_exp.txt"),
            mime="text/plain",
            key="download_uploaded_exp",
        )

        # Run heuristic T-SQL analysis for Snowflake readiness.
        findings = analyze_tsql(exp_content)
        st.subheader("Snowflake Migration Findings")
        if findings:
            for issue in findings:
                severity = issue.severity.capitalize()
                st.markdown(f"**Line {issue.line} – {severity}:** {issue.message}")
                if issue.context:
                    st.code(issue.context, language="sql")
                if issue.recommendation:
                    st.caption(issue.recommendation)
        else:
            st.info("No Snowflake compatibility issues detected by the current heuristic checks.")

        transformed_sql, transform_notes = transform_tsql_to_snow(exp_content)
        st.subheader("Proposed SnowSQL")
        st.text_area("SnowSQL candidate", value=transformed_sql, height=220, key="exp_snow_view")
        st.download_button(
            label="Download SnowSQL candidate",
            data=transformed_sql,
            file_name="exp_translated_snowsql.txt",
            mime="text/plain",
            key="download_snowsql_candidate",
        )
        if transform_notes:
            st.markdown("**Transform Notes**")
            for note in transform_notes:
                st.markdown(f"- {note}")

    # Load comparison definitions and surface any JSON parsing issues.
    comparisons, errors = load_jsonl(comparisons_path)
    if errors:
        with st.expander("comparisons.jsonl issues"):
            for err in errors:
                st.error(err)

    # Exit early if there are no comparisons to work on.
    if not comparisons:
        st.warning("Load or create comparison records to begin.")
        return

    # Let the user select which stored procedure / execution hash to adjust.
    record = pick_record(comparisons)
    if not record:
        return

    # Display the raw record for quick inspection while editing.
    with st.expander("Record details", expanded=True):
        st.json(record.raw)

    current_join_columns = record.join_columns or []
    join_missing = record.is_join_columns_missing()
    if join_missing:
        st.warning("join_columns are missing or set to 0 for the selected record.")
    else:
        st.success(f"join_columns already set: {', '.join(current_join_columns)}")

    # Allow the analyst to adjust join_columns manually, optionally using suggestions.
    st.subheader("Manage join columns")
    suggested_columns = st.session_state.get("candidate_columns", [])
    if suggested_columns:
        st.info(f"Suggested columns detected: {', '.join(suggested_columns)}")

    columns_input = st.text_input(
        "Join columns (comma separated)",
        value=", ".join(current_join_columns),
    )
    columns_list = [col.strip() for col in columns_input.split(",") if col.strip()]
    if st.button("Update join_columns"):
        update_join_columns(comparisons_path, comparisons, record, columns_list)
        st.success(f"Updated join_columns for {record.stored_procedure or 'selected record'}.")

    # Editable view of the comparison filter used by both commands.
    st.subheader("Comparison filter file")
    render_filter_editor("comparison_filter.txt", comparison_filter_path)

    # Trigger the find_keys helper action and stash its output.
    st.subheader("Find Keys")
    st.caption("Use an execution hash different from the one tied to the stored procedure when join_columns are missing.")
    if st.button("Run Find Keys"):
        code, stdout, stderr, command = run_datacompare(
            "find_keys",
            datacompare_path,
            config_path,
            comparison_filter_path,
            workdir_path,
        )
        st.info(f"Command: {' '.join(command)}")
        st.write(f"Exit code: {code}")
        render_command_output(stdout, stderr)
        st.session_state["last_find_stdout"] = stdout
        st.session_state["last_find_stderr"] = stderr
        st.session_state["candidate_columns"] = suggest_columns(stdout)

    st.subheader("Compare")
    if st.button("Run Compare"):
        code, stdout, stderr, command = run_datacompare(
            "compare",
            datacompare_path,
            config_path,
            comparison_filter_path,
            workdir_path,
        )
        st.info(f"Command: {' '.join(command)}")
        st.write(f"Exit code: {code}")
        render_command_output(stdout, stderr)
        st.session_state["last_compare_stdout"] = stdout
        st.session_state["last_compare_stderr"] = stderr

        if detect_match_success(stdout):
            st.success("Match success reported. Update compare_filter.txt with five new execution hashes.")
        else:
            duplicate_status = detect_duplicates(stdout)
            if duplicate_status is False:
                st.error("Match failure without duplicates. Document the failure and capture artifacts per SOP.")
            elif duplicate_status is True:
                st.warning("Duplicates detected on match values. Adjust join_columns and rerun compare.")
            else:
                st.warning("Match failure detected. Review output for next steps.")

    # Provide an editor to maintain the compare_filter list after matches succeed.
    st.subheader("compare_filter.txt")
    render_filter_editor("compare_filter.txt", compare_filter_path)

    if "last_compare_stdout" in st.session_state:
        # Offer the compare stdout for download so it can be attached to analysis tasks.
        st.subheader("Download Compare Output")
        compare_output = st.session_state["last_compare_stdout"]
        st.download_button(
            label="Download compare stdout",
            data=compare_output,
            file_name="compare_stdout.txt",
            mime="text/plain",
        )

    if "last_find_stdout" in st.session_state:
        # Provide the find_keys output in case it needs to be shared or archived.
        st.subheader("Download Find Keys Output")
        st.download_button(
            label="Download find_keys stdout",
            data=st.session_state["last_find_stdout"],
            file_name="find_keys_stdout.txt",
            mime="text/plain",
        )


if __name__ == "__main__":
    # Allow running directly via `streamlit run expanalysistool.app.py`.
    main()
