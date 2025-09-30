"""Streamlit app: auto-detect location from IP and display 7-day weather forecast."""

from __future__ import annotations

from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import requests
import streamlit as st

IP_LOOKUP_URL = "https://ipapi.co/json/"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
DAILY_FIELDS = [
    "weathercode",
    "temperature_2m_max",
    "temperature_2m_min",
    "precipitation_probability_mean",
]

WEATHER_CODE_MAP: Dict[int, str] = {
    0: "Clear sky",
    1: "Mainly clear",
    2: "Partly cloudy",
    3: "Overcast",
    45: "Fog",
    48: "Depositing rime fog",
    51: "Light drizzle",
    53: "Moderate drizzle",
    55: "Dense drizzle",
    56: "Light freezing drizzle",
    57: "Dense freezing drizzle",
    61: "Slight rain",
    63: "Moderate rain",
    65: "Heavy rain",
    66: "Light freezing rain",
    67: "Heavy freezing rain",
    71: "Slight snow fall",
    73: "Moderate snow fall",
    75: "Heavy snow fall",
    77: "Snow grains",
    80: "Slight rain showers",
    81: "Moderate rain showers",
    82: "Violent rain showers",
    85: "Slight snow showers",
    86: "Heavy snow showers",
    95: "Thunderstorm",
    96: "Thunderstorm with slight hail",
    99: "Thunderstorm with heavy hail",
}

st.set_page_config(
    page_title="Weather Near Me",
    page_icon="⛅",
    layout="centered",
)


@st.cache_data(show_spinner="Looking up your location…")
def lookup_location() -> Dict[str, Optional[str]]:
    response = requests.get(IP_LOOKUP_URL, timeout=5)
    response.raise_for_status()
    payload = response.json()

    latitude = payload.get("latitude")
    longitude = payload.get("longitude")
    if latitude is None or longitude is None:
        raise ValueError("IP lookup did not return latitude/longitude.")

    return {
        "ip": payload.get("ip"),
        "city": payload.get("city"),
        "region": payload.get("region"),
        "country": payload.get("country_name"),
        "latitude": float(latitude),
        "longitude": float(longitude),
    }


@st.cache_data(show_spinner="Fetching 7-day forecast…")
def fetch_forecast(latitude: float, longitude: float) -> Dict[str, List]:
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "daily": ",".join(DAILY_FIELDS),
        "timezone": "auto",
    }
    response = requests.get(FORECAST_URL, params=params, timeout=10)
    response.raise_for_status()
    data = response.json()
    if "daily" not in data:
        raise ValueError("Forecast response missing daily data.")
    return data


def describe_weather(code: Optional[int]) -> str:
    if code is None:
        return "—"
    return WEATHER_CODE_MAP.get(int(code), f"Weather code {code}")


def to_unit(temp_c: Optional[float], unit: str) -> Optional[float]:
    if temp_c is None:
        return None
    if unit == "Fahrenheit":
        return round(temp_c * 9.0 / 5.0 + 32.0, 1)
    return round(temp_c, 1)


def build_forecast_frame(raw: Dict[str, Dict[str, List]], unit: str) -> pd.DataFrame:
    daily = raw.get("daily", {})
    dates = daily.get("time", [])
    codes = daily.get("weathercode", [])
    highs = daily.get("temperature_2m_max", [])
    lows = daily.get("temperature_2m_min", [])
    precip = daily.get("precipitation_probability_mean", [])

    rows: List[Dict[str, Optional[str]]] = []
    for idx, date_str in enumerate(dates[:7]):
        try:
            dt = datetime.fromisoformat(date_str)
            label = dt.strftime("%a %b %d")
        except ValueError:
            label = date_str

        code = codes[idx] if idx < len(codes) else None
        high_c = highs[idx] if idx < len(highs) else None
        low_c = lows[idx] if idx < len(lows) else None
        precip_pct = precip[idx] if idx < len(precip) else None

        rows.append(
            {
                "Date": label,
                "Conditions": describe_weather(code),
                f"High (°{'F' if unit == 'Fahrenheit' else 'C'})": to_unit(high_c, unit),
                f"Low (°{'F' if unit == 'Fahrenheit' else 'C'})": to_unit(low_c, unit),
                "Precip (%)": None if precip_pct is None else int(round(precip_pct)),
            }
        )

    return pd.DataFrame(rows)


def main() -> None:
    st.title("Next 7 Days Near You")
    st.caption("Powered by Open-Meteo and automatic IP geolocation.")

    location_error: Optional[str] = None
    location: Optional[Dict[str, Optional[str]]] = None

    try:
        location = lookup_location()
    except Exception as exc:  # pylint: disable=broad-except
        location_error = str(exc)

    col1, col2 = st.columns([3, 2])
    with col1:
        st.subheader("Where you are")
        if location:
            city = location.get("city") or "Unknown city"
            pieces = [city]
            if location.get("region"):
                pieces.append(location["region"])
            if location.get("country"):
                pieces.append(location["country"])
            st.success(
                ", ".join(pieces) +
                f" • Lat {location['latitude']:.2f}, Lon {location['longitude']:.2f}"
            )
            if location.get("ip"):
                st.caption(f"Detected from IP {location['ip']}")
        else:
            st.error("Automatic location lookup failed.")
            if location_error:
                st.caption(location_error)

    with col2:
        st.subheader("Adjust coordinates")
        default_lat = location["latitude"] if location else 0.0
        default_lon = location["longitude"] if location else 0.0
        latitude = st.number_input("Latitude", value=float(default_lat), format="%.4f")
        longitude = st.number_input("Longitude", value=float(default_lon), format="%.4f")

    unit = st.radio("Temperature units", options=("Celsius", "Fahrenheit"), horizontal=True)

    if not latitude and not longitude:
        st.info("Set latitude and longitude to fetch the forecast.")
        return

    try:
        forecast = fetch_forecast(latitude, longitude)
    except Exception as exc:  # pylint: disable=broad-except
        st.error("Could not retrieve weather data.")
        st.caption(str(exc))
        return

    frame = build_forecast_frame(forecast, unit)
    if frame.empty:
        st.warning("No forecast data returned. Try different coordinates.")
        return

    st.subheader("7-day outlook")
    st.dataframe(frame, use_container_width=True)

    chart_data = frame.set_index("Date")[[col for col in frame.columns if col.startswith("High") or col.startswith("Low")]]
    st.line_chart(chart_data)

    st.caption("Data from https://open-meteo.com with precipitation shown as daily mean probability.")


if __name__ == "__main__":
    main()
