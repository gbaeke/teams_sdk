from __future__ import annotations

from typing import Any
from urllib.parse import urlencode

import httpx

GEOCODING_URL = "https://geocoding-api.open-meteo.com/v1/search"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
CURRENT_FIELDS = (
    "temperature_2m,relative_humidity_2m,apparent_temperature,"
    "precipitation,weather_code,wind_speed_10m"
)

WEATHER_CODES = {
    0: "clear sky",
    1: "mainly clear",
    2: "partly cloudy",
    3: "overcast",
    45: "fog",
    48: "depositing rime fog",
    51: "light drizzle",
    53: "moderate drizzle",
    55: "dense drizzle",
    56: "light freezing drizzle",
    57: "dense freezing drizzle",
    61: "slight rain",
    63: "moderate rain",
    65: "heavy rain",
    66: "light freezing rain",
    67: "heavy freezing rain",
    71: "slight snow fall",
    73: "moderate snow fall",
    75: "heavy snow fall",
    77: "snow grains",
    80: "slight rain showers",
    81: "moderate rain showers",
    82: "violent rain showers",
    85: "slight snow showers",
    86: "heavy snow showers",
    95: "thunderstorm",
    96: "thunderstorm with slight hail",
    99: "thunderstorm with heavy hail",
}


async def get_current_weather(location: str) -> dict[str, Any]:
    """Return current weather for a human-readable location."""
    location = location.strip()
    if not location:
        return {"error": "A location is required."}

    async with httpx.AsyncClient(timeout=10) as http:
        geocoded = await _geocode_location(http, location)
        if "error" in geocoded:
            return geocoded

        response = await http.get(
            FORECAST_URL,
            params={
                "latitude": geocoded["latitude"],
                "longitude": geocoded["longitude"],
                "current": CURRENT_FIELDS,
                "timezone": "auto",
            },
        )
        response.raise_for_status()
        data = response.json()

    current = data.get("current")
    units = data.get("current_units", {})
    if not current:
        return {"error": f"No current weather was returned for {location}."}

    weather_code = current.get("weather_code")
    return {
        "location": geocoded,
        "time": current.get("time"),
        "temperature": _with_unit(current.get("temperature_2m"), units.get("temperature_2m")),
        "apparent_temperature": _with_unit(
            current.get("apparent_temperature"),
            units.get("apparent_temperature"),
        ),
        "relative_humidity": _with_unit(
            current.get("relative_humidity_2m"),
            units.get("relative_humidity_2m"),
        ),
        "precipitation": _with_unit(current.get("precipitation"), units.get("precipitation")),
        "wind_speed": _with_unit(current.get("wind_speed_10m"), units.get("wind_speed_10m")),
        "weather_code": weather_code,
        "conditions": WEATHER_CODES.get(weather_code, "unknown conditions"),
        "source": "Open-Meteo",
        "source_url": _forecast_url(geocoded),
    }


async def _geocode_location(http: httpx.AsyncClient, location: str) -> dict[str, Any]:
    response = await http.get(
        GEOCODING_URL,
        params={
            "name": location,
            "count": 1,
            "language": "en",
            "format": "json",
        },
    )
    response.raise_for_status()
    results = response.json().get("results", [])
    if not results:
        return {"error": f"Could not find weather coordinates for '{location}'."}

    match = results[0]
    return {
        "name": match.get("name"),
        "country": match.get("country"),
        "admin1": match.get("admin1"),
        "latitude": match.get("latitude"),
        "longitude": match.get("longitude"),
        "timezone": match.get("timezone"),
    }


def _with_unit(value: Any, unit: str | None) -> str | None:
    if value is None:
        return None
    return f"{value} {unit}" if unit else str(value)


def _forecast_url(location: dict[str, Any]) -> str:
    query = urlencode(
        {
            "latitude": location["latitude"],
            "longitude": location["longitude"],
            "current": CURRENT_FIELDS,
            "timezone": "auto",
        }
    )
    return f"{FORECAST_URL}?{query}"
