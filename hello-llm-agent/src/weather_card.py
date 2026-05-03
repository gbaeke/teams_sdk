from __future__ import annotations

from typing import Any, Literal, cast

from microsoft_teams.api import MessageActivityInput
from microsoft_teams.cards import (
    AdaptiveCard,
    Column,
    ColumnSet,
    Fact,
    FactSet,
    OpenUrlAction,
    TextBlock,
)

TextColor = Literal["Default", "Dark", "Light", "Accent", "Good", "Warning", "Attention"]


def create_weather_card(weather: dict[str, Any]) -> MessageActivityInput:
    location = weather.get("location", {})
    location_name = _location_name(location)
    conditions = str(weather.get("conditions") or "Current conditions")
    temperature = str(weather.get("temperature") or "n/a")
    apparent_temperature = str(weather.get("apparent_temperature") or "n/a")
    source = str(weather.get("source") or "Open-Meteo")
    updated = str(weather.get("time") or "unknown time")
    source_url = str(weather.get("source_url") or "")
    icon = _weather_icon(weather.get("weather_code"))
    accent = _weather_accent(weather.get("weather_code"))

    card = AdaptiveCard(
        version="1.5",
        schema="http://adaptivecards.io/schemas/adaptive-card.json",
        fallback_text=f"{location_name}: {temperature}, {conditions}. Source: {source}.",
        body=[
            ColumnSet(
                columns=[
                    Column(
                        width="stretch",
                        items=cast(Any, [
                            TextBlock(
                                text=location_name,
                                size="Medium",
                                weight="Bolder",
                                wrap=True,
                            ),
                            TextBlock(
                                text=conditions.title(),
                                color=accent,
                                spacing="None",
                                wrap=True,
                            ),
                        ]),
                    ),
                    Column(
                        width="auto",
                        vertical_content_alignment="Center",
                        items=cast(Any, [
                            TextBlock(
                                text=icon,
                                size="ExtraLarge",
                                horizontal_alignment="Right",
                                wrap=False,
                            )
                        ]),
                    ),
                ],
            ),
            TextBlock(
                text=temperature,
                size="ExtraLarge",
                weight="Bolder",
                spacing="Medium",
                wrap=True,
            ),
            TextBlock(
                text=f"Feels like {apparent_temperature}",
                is_subtle=True,
                spacing="None",
                wrap=True,
            ),
            FactSet(
                spacing="Medium",
                facts=[
                    Fact(title="Humidity", value=str(weather.get("relative_humidity") or "n/a")),
                    Fact(title="Wind", value=str(weather.get("wind_speed") or "n/a")),
                    Fact(title="Precipitation", value=str(weather.get("precipitation") or "n/a")),
                ],
            ),
            TextBlock(
                text=f"Source: {source} | Updated: {updated}",
                size="Small",
                is_subtle=True,
                spacing="Medium",
                wrap=True,
            ),
        ],
        actions=[
            OpenUrlAction(
                title="Open Open-Meteo data",
                url=source_url,
            )
        ]
        if source_url
        else None,
    )

    return MessageActivityInput(summary=f"Weather for {location_name}").add_card(card)


def _location_name(location: Any) -> str:
    if not isinstance(location, dict):
        return "Weather"

    parts = [
        location.get("name"),
        location.get("admin1"),
        location.get("country"),
    ]
    return ", ".join(str(part) for part in parts if part)


def _weather_icon(weather_code: Any) -> str:
    if weather_code == 0:
        return "☀"
    if weather_code in {1, 2}:
        return "⛅"
    if weather_code == 3:
        return "☁"
    if weather_code in {45, 48}:
        return "🌫"
    if weather_code in {51, 53, 55, 56, 57, 61, 63, 65, 66, 67, 80, 81, 82}:
        return "🌧"
    if weather_code in {71, 73, 75, 77, 85, 86}:
        return "❄"
    if weather_code in {95, 96, 99}:
        return "⛈"
    return "🌡"


def _weather_accent(weather_code: Any) -> TextColor:
    if weather_code in {0, 1}:
        return "Good"
    if weather_code in {51, 53, 55, 61, 63, 65, 80, 81, 82, 95, 96, 99}:
        return "Warning"
    if weather_code in {66, 67, 71, 73, 75, 77, 85, 86}:
        return "Accent"
    return "Default"
