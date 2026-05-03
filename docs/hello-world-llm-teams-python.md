# Hello World LLM App with the Microsoft Teams SDK for Python

This guide describes the smallest practical path for creating a Python Microsoft Teams app that replies to user messages with an LLM-generated answer.

Last researched: 2026-05-03.

## Current SDK State

Microsoft's Python support for the Microsoft Teams SDK was announced as generally available on 2026-05-01. The core package for a Python Teams agent is `microsoft-teams-apps`, imported as `microsoft_teams.apps`.

The SDK also has AI abstractions in `microsoft-teams-ai`, but the `microsoft-teams-openai` package is marked deprecated on PyPI and recommends using the official OpenAI Python SDK instead. For a hello-world LLM app, use the Teams SDK for Teams message handling and the OpenAI SDK for model calls.

## Prerequisites

Install these locally:

1. Python 3.12 or newer.
2. Node.js/npm, for the Teams CLI.
3. Teams CLI:

   ```bash
   npm install -g @microsoft/teams.cli@preview
   teams --version
   ```

4. A Microsoft 365 account where custom app upload, also called sideloading, is enabled.
5. A public HTTPS tunnel to your local app, such as Microsoft Dev Tunnels or ngrok.
6. An OpenAI API key in `OPENAI_API_KEY`.

## 1. Create the Project

The Teams SDK quickstart scaffolds a basic echo agent:

```bash
teams project new python hello-llm-agent --template echo
cd hello-llm-agent
```

The generated project includes:

- `src/main.py`: the Python entry point.
- `appPackage/manifest.json`: the Teams app manifest used for sideloading.
- icon placeholders under `appPackage/`.

## 2. Add LLM Dependencies

Install the Teams app framework and the OpenAI Python SDK:

```bash
pip install microsoft-teams-apps openai python-dotenv
```

If the scaffold created a `pyproject.toml`, add the same packages there instead so installs are repeatable.

Create a local `.env` file:

```bash
OPENAI_API_KEY=replace-with-your-key
OPENAI_MODEL=gpt-5.2
```

Do not commit `.env`.

## 3. Implement the LLM Message Handler

Replace the echo handler in `src/main.py` with an LLM-backed handler like this:

```python
import asyncio
import os

from dotenv import load_dotenv
from openai import AsyncOpenAI

from microsoft_teams.api import MessageActivity
from microsoft_teams.apps import ActivityContext, App

load_dotenv()

app = App()
client = AsyncOpenAI()
model = os.getenv("OPENAI_MODEL", "gpt-5.2")


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]):
    user_text = (ctx.activity.text or "").strip()

    if not user_text:
        await ctx.send("Send me a message and I will answer with an LLM.")
        return

    response = await client.responses.create(
        model=model,
        instructions=(
            "You are a concise assistant running inside Microsoft Teams. "
            "Answer clearly and keep responses short unless the user asks for detail."
        ),
        input=user_text,
    )

    await ctx.send(response.output_text)


if __name__ == "__main__":
    asyncio.run(app.start())
```

This keeps the Teams-specific responsibility in the Teams SDK: receive a Teams activity and send a reply. The LLM responsibility stays in the OpenAI SDK: call the model and return generated text.

## 4. Run Locally

Start the app:

```bash
python src/main.py
```

The Teams SDK default HTTP server listens on port `3978`. Keep this process running while testing.

## 5. Expose the Local Server

Teams needs a public HTTPS endpoint for your local bot. Start a tunnel to port `3978`.

Example with Dev Tunnels:

```bash
devtunnel host -p 3978 --allow-anonymous
```

Your Teams messaging endpoint will be:

```text
https://<your-tunnel-host>/api/messages
```

## 6. Register the App with Teams

Sign in and verify sideloading:

```bash
teams login
teams status
```

`teams status` should show sideloading enabled. Then register the bot infrastructure:

```bash
teams app create \
  --name hello-llm-agent \
  --endpoint https://<your-tunnel-host>/api/messages \
  --env .env
```

This creates a Teams-managed bot by default and writes Teams credentials such as `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` to `.env`.

## 7. Install and Test in Teams

Use the install link printed by `teams app create`, or retrieve it later:

```bash
teams app list
teams app get <teamsAppId> --install-link
```

Open the install link while signed in to Teams, add the app, and send it a message such as:

```text
hello, summarize what you can do in one sentence
```

Expected behavior: Teams sends the message to your local Python app through the tunnel, the Python app calls the LLM, and the generated answer is posted back into the Teams conversation.

## 8. Troubleshooting Checklist

- `teams status` does not show sideloading enabled: ask the tenant admin to enable custom app upload.
- No messages arrive locally: verify the tunnel forwards to port `3978` and the endpoint ends with `/api/messages`.
- Authentication errors from Teams: rerun `teams app create` or `teams app doctor`.
- OpenAI call fails: verify `OPENAI_API_KEY` is set in the same shell where `python src/main.py` runs.
- Model name fails: set `OPENAI_MODEL` to a model available to your OpenAI project.

## Production Notes

For production, replace the local tunnel with a deployed HTTPS endpoint. The Teams SDK documentation calls out Azure App Service, Azure Container Apps, and similar compute options as suitable hosting targets. Keep secrets in a managed secret store rather than `.env`, add logging/telemetry, and consider rate limits and content safety before exposing the bot broadly.

## Sources

- Microsoft Learn overview: <https://learn.microsoft.com/en-us/python/api/msteams-sdk-python/overview?view=msteams-sdk-python-latest>
- Teams SDK Python quickstart: <https://microsoft.github.io/teams-sdk/python/getting-started/quickstart/>
- Teams SDK Python running in Teams guide: <https://microsoft.github.io/teams-sdk/python/getting-started/running-in-teams/>
- Teams SDK core concepts: <https://microsoft.github.io/teams-sdk/teams/core-concepts/>
- Microsoft 365 Developer Blog GA announcement: <https://devblogs.microsoft.com/microsoft365dev/python-support-for-the-microsoft-teams-sdk-is-now-generally-available/>
- `microsoft-teams-apps` on PyPI: <https://pypi.org/project/microsoft-teams-apps/>
- `microsoft-teams-openai` deprecation note on PyPI: <https://pypi.org/project/microsoft-teams-openai/>
- OpenAI Python quickstart: <https://platform.openai.com/docs/overview?lang=python>
