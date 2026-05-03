import asyncio
import json
import os
from pathlib import Path
from typing import cast

import httpx
from dotenv import load_dotenv
from openai import AsyncOpenAI
from openai.types.responses import (
    FunctionToolParam,
    ResponseFunctionToolCall,
    ResponseInputParam,
)

from microsoft_teams.api import MessageActivity
from microsoft_teams.apps import ActivityContext, App
from src.weather_card import create_weather_card
from src.weather import get_current_weather

APP_DIR = Path(__file__).resolve().parents[1]
REPO_DIR = APP_DIR.parent

load_dotenv(REPO_DIR / ".env")
load_dotenv(APP_DIR / ".env", override=True)

app = App()
client = AsyncOpenAI()
model = os.getenv("OPENAI_MODEL", "gpt-5.2")

WEATHER_TOOL: FunctionToolParam = {
    "type": "function",
    "name": "get_current_weather",
    "description": "Get actual current weather for a city or place using Open-Meteo.",
    "parameters": {
        "type": "object",
        "properties": {
            "location": {
                "type": "string",
                "description": "City or place name, optionally with country or region.",
            }
        },
        "required": ["location"],
        "additionalProperties": False,
    },
    "strict": True,
}


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Answer Teams messages with an LLM-generated response."""
    user_text = (ctx.activity.text or "").strip()

    if not user_text:
        await ctx.send("Send me a message and I will answer with an LLM.")
        return

    response = await client.responses.create(
        model=model,
        instructions=(
            "You are a concise assistant running inside Microsoft Teams. "
            "Answer clearly and keep responses short unless the user asks for detail. "
            "When the user asks for current weather, use the get_current_weather tool."
        ),
        tools=[WEATHER_TOOL],
        input=user_text,
    )

    tool_calls = [
        item
        for item in response.output
        if isinstance(item, ResponseFunctionToolCall) and item.name == "get_current_weather"
    ]
    if tool_calls:
        weather_results = [await _run_weather_tool(call) for call in tool_calls]
        weather_result = weather_results[0]
        if "error" not in weather_result:
            await ctx.send(create_weather_card(weather_result))
            return

        follow_up_input = cast(
            ResponseInputParam,
            [
                *[
                    {
                        "type": "function_call_output",
                        "call_id": call.call_id,
                        "output": json.dumps(result),
                    }
                    for call, result in zip(tool_calls, weather_results, strict=True)
                ],
            ],
        )
        response = await client.responses.create(
            model=model,
            instructions=(
                "Use the weather tool result to answer the user. "
                "Mention the location, conditions, temperature, and source."
            ),
            previous_response_id=response.id,
            input=follow_up_input,
        )

    await ctx.send(response.output_text)


async def _run_weather_tool(call: ResponseFunctionToolCall) -> dict[str, object]:
    try:
        arguments = json.loads(call.arguments)
        location = arguments.get("location", "")
        return await get_current_weather(str(location))
    except (httpx.HTTPError, json.JSONDecodeError) as exc:
        return {"error": f"Weather lookup failed: {exc}"}


def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()
