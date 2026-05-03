import asyncio
import json
import os
from pathlib import Path
from typing import Any, cast

import httpx
from dotenv import load_dotenv
from openai import AsyncOpenAI, BadRequestError, NotFoundError
from openai.types.responses import (
    FunctionToolParam,
    Response as OpenAIResponse,
    ResponseFunctionToolCall,
    ResponseInputParam,
    WebSearchToolParam,
)
from openai.types.responses.response_output_text import AnnotationURLCitation
from openai.types.responses.response_output_message import ResponseOutputMessage

from microsoft_teams.api import (
    CardAction,
    CardActionType,
    MessageActivity,
    MessageActivityInput,
    SuggestedActions,
)
from microsoft_teams.apps import ActivityContext, App
from src.weather_card import create_weather_card
from src.weather import get_current_weather

APP_DIR = Path(__file__).resolve().parents[1]
REPO_DIR = APP_DIR.parent

load_dotenv(REPO_DIR / ".env")
load_dotenv(APP_DIR / ".env", override=True)

app = App()
app.tab("about", str(APP_DIR / "static" / "about"))
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

WEB_SEARCH_TOOL: WebSearchToolParam = {"type": "web_search"}

SUGGESTED_PROMPTS = [
    CardAction(
        type=CardActionType.IM_BACK,
        title="Weather in Paris",
        value="weather in Paris",
    ),
    CardAction(
        type=CardActionType.IM_BACK,
        title="Latest AI news",
        value="search the web for the latest AI news today",
    ),
    CardAction(
        type=CardActionType.IM_BACK,
        title="Who am I?",
        value="who am I?",
    ),
    CardAction(
        type=CardActionType.IM_BACK,
        title="Explain this bot",
        value="explain what this bot can do",
    ),
]


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Answer Teams messages with an LLM-generated response."""
    # Strip @mentions so commands picked from the Teams command list work in
    # team/group scope, where the bot's @mention is prepended to text.
    ctx.activity.strip_mentions_text()
    user_text = (ctx.activity.text or "").strip()
    ctx.logger.info("incoming text=%r", user_text)

    if not user_text:
        await _send_ai_text(ctx, "Send me a message and I will answer with an LLM.")
        return

    if _is_forget_command(user_text):
        await ctx.storage.async_delete(_memory_key(ctx))
        await _send_ai_text(ctx, "OK, I've cleared our conversation history.")
        return

    if _is_who_am_i_question(user_text):
        await _handle_who_am_i_from_teams(ctx)
        return

    memory_key = _memory_key(ctx)
    previous_id = await ctx.storage.async_get(memory_key)
    ctx.logger.info("memory load key=%s previous_id=%s", memory_key, previous_id)

    primary_instructions = (
        "You are a concise assistant running inside Microsoft Teams. "
        "Answer clearly and keep responses short unless the user asks for detail. "
        "When the user asks for current weather, use the get_current_weather tool. "
        "When the user asks about current events, recent news, or anything that "
        "needs up-to-date information from the web, use the web_search tool."
    )

    try:
        response = await _stream_llm_text(
            ctx,
            instructions=primary_instructions,
            input=cast(ResponseInputParam, user_text),
            tools=[WEATHER_TOOL, WEB_SEARCH_TOOL],
            previous_response_id=previous_id,
        )
    except (NotFoundError, BadRequestError):
        # Stored response is unusable — expired (≈30 day TTL) or has
        # unfulfilled tool calls. Drop it and start a fresh chain.
        await ctx.storage.async_delete(memory_key)
        response = await _stream_llm_text(
            ctx,
            instructions=primary_instructions,
            input=cast(ResponseInputParam, user_text),
            tools=[WEATHER_TOOL, WEB_SEARCH_TOOL],
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
            completion = await client.responses.create(
                model=model,
                instructions=(
                    "A weather card was rendered for the user. "
                    "Reply with a single short acknowledgement; the user does not see this text."
                ),
                previous_response_id=response.id,
                input=cast(
                    ResponseInputParam,
                    [
                        {
                            "type": "function_call_output",
                            "call_id": call.call_id,
                            "output": json.dumps(result),
                        }
                        for call, result in zip(tool_calls, weather_results, strict=True)
                    ],
                ),
                max_output_tokens=32,
            )
            await ctx.storage.async_set(memory_key, completion.id)
            ctx.logger.info(
                "memory save (post-card) key=%s response_id=%s", memory_key, completion.id
            )
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
        response = await _stream_llm_text(
            ctx,
            instructions=(
                "Use the weather tool result to answer the user. "
                "Mention the location, conditions, temperature, and source."
            ),
            input=follow_up_input,
            previous_response_id=response.id,
        )

    await ctx.storage.async_set(memory_key, response.id)
    ctx.logger.info("memory save key=%s response_id=%s", memory_key, response.id)

    citations = _extract_url_citations(response)
    if citations:
        footer_lines = ["", "", "**Sources:**"]
        footer_lines.extend(f"- [{c.title or c.url}]({c.url})" for c in citations)
        ctx.stream.emit("\n".join(footer_lines))

    await _finalize_ai_stream(ctx)


async def _stream_llm_text(
    ctx: ActivityContext[MessageActivity],
    *,
    instructions: str,
    input: ResponseInputParam,
    tools: list[FunctionToolParam] | None = None,
    previous_response_id: str | None = None,
) -> OpenAIResponse:
    """Run the Responses API with streaming and forward text deltas to the Teams stream."""
    kwargs: dict[str, Any] = {
        "model": model,
        "instructions": instructions,
        "input": input,
        "stream": True,
    }
    if tools is not None:
        kwargs["tools"] = tools
    if previous_response_id is not None:
        kwargs["previous_response_id"] = previous_response_id

    response_stream = await client.responses.create(**kwargs)

    final_response: OpenAIResponse | None = None
    async for event in response_stream:
        if event.type == "response.output_text.delta":
            if event.delta:
                ctx.stream.emit(event.delta)
        elif event.type == "response.web_search_call.in_progress":
            ctx.stream.update("Searching the web…")
        elif event.type == "response.web_search_call.searching":
            ctx.stream.update("Reading web results…")
        elif event.type == "response.completed":
            final_response = event.response

    if final_response is None:
        raise RuntimeError("Response stream ended without a completion event.")
    return final_response


def _extract_url_citations(response: OpenAIResponse) -> list[AnnotationURLCitation]:
    """Pull URL citations out of a Responses API response (de-duplicated by URL)."""
    seen: dict[str, AnnotationURLCitation] = {}
    for item in response.output:
        if not isinstance(item, ResponseOutputMessage):
            continue
        for part in item.content:
            for annotation in getattr(part, "annotations", []) or []:
                if isinstance(annotation, AnnotationURLCitation) and annotation.url not in seen:
                    seen[annotation.url] = annotation
    return list(seen.values())


async def _finalize_ai_stream(ctx: ActivityContext[MessageActivity]) -> None:
    """Attach the AI-generated label and suggested prompts, then close the stream."""
    final = (
        MessageActivityInput()
        .add_ai_generated()
        .with_suggested_actions(
            SuggestedActions(
                to=[ctx.activity.from_.id],
                actions=SUGGESTED_PROMPTS,
            )
        )
    )
    ctx.stream.emit(final)
    await ctx.stream.close()


async def _run_weather_tool(call: ResponseFunctionToolCall) -> dict[str, object]:
    try:
        arguments = json.loads(call.arguments)
        location = arguments.get("location", "")
        return await get_current_weather(str(location))
    except (httpx.HTTPError, json.JSONDecodeError) as exc:
        return {"error": f"Weather lookup failed: {exc}"}


async def _send_ai_text(ctx: ActivityContext[MessageActivity], text: str) -> None:
    message = (
        MessageActivityInput(text=text)
        .add_ai_generated()
        .with_suggested_actions(
            SuggestedActions(
                to=[ctx.activity.from_.id],
                actions=SUGGESTED_PROMPTS,
            )
        )
    )
    await ctx.send(message)


def _memory_key(ctx: ActivityContext[MessageActivity]) -> str:
    sender = ctx.activity.from_
    conversation = ctx.activity.conversation
    user_id = sender.aad_object_id or sender.id or "unknown_user"
    conv_id = conversation.id or "unknown_conversation"
    return f"prev_resp:{user_id}:{conv_id}"


def _is_forget_command(text: str) -> bool:
    normalized = text.strip().lower().rstrip("?!.")
    return normalized in {
        "/forget",
        "/reset",
        "forget",
        "reset",
        # Also match the manifest command description, in case Teams sends
        # the description text instead of the title from the command picker.
        "clear conversation memory",
        "clear memory",
    }


def _is_who_am_i_question(text: str) -> bool:
    normalized = text.strip().lower().rstrip("?!.")
    return normalized in {
        "who am i",
        "whoami",
        "/whoami",
        "what is my name",
        "what's my name",
        # Manifest command description fallback.
        "show what teams knows about you",
    }


async def _handle_who_am_i_from_teams(ctx: ActivityContext[MessageActivity]) -> None:
    sender = ctx.activity.from_
    conversation = ctx.activity.conversation

    name = sender.name or "unknown name"
    aad_object_id = sender.aad_object_id or "not provided"
    teams_user_id = sender.id or "not provided"
    tenant_id = conversation.tenant_id or "not provided"
    conversation_type = conversation.conversation_type or "not provided"

    await _send_ai_text(
        ctx,
        (
            f"Teams says you are **{name}**.\n\n"
            f"- Teams user ID: `{teams_user_id}`\n"
            f"- Entra object ID: `{aad_object_id}`\n"
            f"- Tenant ID: `{tenant_id}`\n"
            f"- Conversation type: `{conversation_type}`"
        ),
    )


def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()
