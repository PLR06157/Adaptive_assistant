# Quickstart: Create and Test a Basic Agent

This guide walks through creating a custom engine agent that echoes back messages, with implementations in Python, JavaScript, and C#/.NET.

---

## Python

### Prerequisites
- Python 3.9 or newer ([download](https://www.python.org/downloads/))
- Visual Studio Code with Python extension (recommended)

### Step 1: Initialize Project and Install SDK

```bash
mkdir echo
cd echo
code .
```

In VS Code, create a virtual environment:
- Press `F1` > `Python: Create environment` > Select `Venv` > Select Python installation

Install the SDK:
```bash
pip install microsoft-agents-hosting-aiohttp
```

### Step 2: Create `start_server.py`

```python
from os import environ
from microsoft_agents.hosting.core import AgentApplication, AgentAuthConfiguration
from microsoft_agents.hosting.aiohttp import (
   start_agent_process,
   jwt_authorization_middleware,
   CloudAdapter,
)
from aiohttp.web import Request, Response, Application, run_app


def start_server(
   agent_application: AgentApplication, auth_configuration: AgentAuthConfiguration
):
   async def entry_point(req: Request) -> Response:
      agent: AgentApplication = req.app["agent_app"]
      adapter: CloudAdapter = req.app["adapter"]
      return await start_agent_process(
            req,
            agent,
            adapter,
      )

   APP = Application(middlewares=[jwt_authorization_middleware])
   APP.router.add_post("/api/messages", entry_point)
   APP.router.add_get("/api/messages", lambda _: Response(status=200))
   APP["agent_configuration"] = auth_configuration
   APP["agent_app"] = agent_application
   APP["adapter"] = agent_application.adapter

   try:
      run_app(APP, host="localhost", port=environ.get("PORT", 3978))
   except Exception as error:
      raise error
```

### Step 3: Create `app.py`

```python
from microsoft_agents.hosting.core import (
   AgentApplication,
   TurnState,
   TurnContext,
   MemoryStorage,
)
from microsoft_agents.hosting.aiohttp import CloudAdapter
from start_server import start_server

# Create AGENT_APP instance
AGENT_APP = AgentApplication[TurnState](
    storage=MemoryStorage(), adapter=CloudAdapter()
)

async def _help(context: TurnContext, _: TurnState):
    await context.send_activity(
        "Welcome to the Echo Agent sample. "
        "Type /help for help or send a message to see the echo feature in action."
    )

AGENT_APP.conversation_update("membersAdded")(_help)

AGENT_APP.message("/help")(_help)


@AGENT_APP.activity("message")
async def on_message(context: TurnContext, _):
    await context.send_activity(f"you said: {context.activity.text}")


if __name__ == "__main__":
    try:
        start_server(AGENT_APP, None)
    except Exception as error:
        raise error
```

### Step 4: Run

```bash
python app.py
```

Expected output:
```
======== Running on http://localhost:3978 ========
(Press CTRL+C to quit)
```

---

## JavaScript / Node.js

### Prerequisites
- Node.js v22 or newer ([download](https://nodejs.org/))
- Visual Studio Code (recommended)

### Step 1: Initialize Project and Install SDK

```bash
mkdir echo
cd echo
npm init -y
npm install @microsoft/agents-hosting-express
code .
```

### Step 2: Create `index.mjs`

```javascript
import { startServer } from '@microsoft/agents-hosting-express'
import { AgentApplication, MemoryStorage } from '@microsoft/agents-hosting'

class EchoAgent extends AgentApplication {
  constructor (storage) {
    super({ storage })

    this.onConversationUpdate('membersAdded', this._help)
    this.onMessage('/help', this._help)
    this.onActivity('message', this._echo)
  }

  _help = async context =>
    await context.sendActivity(`Welcome to the Echo Agent sample.
      Type /help for help or send a message to see the echo feature in action.`)

  _echo = async (context, state) => {
    let counter = state.getValue('conversation.counter') || 0
    await context.sendActivity(`[${counter++}]You said: ${context.activity.text}`)
    state.setValue('conversation.counter', counter)
  }
}

startServer(new EchoAgent(new MemoryStorage()))
```

### Step 3: Run

```bash
node index.mjs
```

Expected output:
```
Server listening to port 3978 on sdk 0.6.18 for appId undefined debug undefined
```

---

## C# / .NET

### Prerequisites
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download)
- Visual Studio or Visual Studio Code
- Download the [QuickStart sample](https://github.com/microsoft/Agents/tree/main/samples/dotnet/quickstart)

### Step 1: Open and Run

Open `QuickStart.csproj` in Visual Studio and run the project. Agent runs on port 3978.

### Step 2: Program.cs Configuration

```csharp
// Add AgentApplicationOptions from appsettings section "AgentApplication"
builder.AddAgentApplicationOptions();

// Add the AgentApplication with custom MyAgent class
builder.AddAgent<MyAgent>();

// Add storage (MemoryStorage for development)
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Route to accept messages
var incomingRoute = app.MapPost("/api/messages",
    async (HttpRequest request, HttpResponse response, IAgentHttpAdapter adapter,
           IAgent agent, CancellationToken cancellationToken) =>
{
    await adapter.ProcessAsync(request, response, agent, cancellationToken);
});
```

### Step 3: MyAgent.cs

```csharp
public MyAgent(AgentApplicationOptions options) : base(options)
{
   OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeMessageAsync);
   OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);
}

private async Task WelcomeMessageAsync(ITurnContext turnContext, ITurnState turnState,
    CancellationToken cancellationToken)
{
   await turnContext.SendActivityAsync("Welcome to the Echo Agent sample.",
       cancellationToken: cancellationToken);
}

private async Task OnMessageAsync(ITurnContext turnContext, ITurnState turnState,
    CancellationToken cancellationToken)
{
   await turnContext.SendActivityAsync($"You said: {turnContext.Activity.Text}",
       cancellationToken: cancellationToken);
}
```

---

## Testing the Agent (All Platforms)

### Install Microsoft 365 Agents Playground

**Python / Node.js (global):**
```bash
npm install -g @microsoft/teams-app-test-tool
teamsapptester
```

**Node.js (local):**
```bash
npm install -D @microsoft/teams-app-test-tool
node_modules/.bin/teamsapptester
```

**.NET (Windows):**
```bash
winget install agentsplayground
agentsplayground
```

### How to Test

1. Keep agent running in terminal 1
2. Run test tool in terminal 2
3. Browser opens to `http://localhost:56150`
4. Send messages to see echo responses
5. Type `/help` for help messages

---

## Key Concepts

### Event Routing

| Event | Handler | Response |
|-------|---------|----------|
| `membersAdded` | `_help` / `WelcomeMessageAsync` | Welcome message |
| `/help` message | `_help` / specific handler | Help message |
| Any `message` | `on_message` / `_echo` / `OnMessageAsync` | Echo user message |

### Storage Options
- **MemoryStorage** — Default for development (state lost on restart)
- **BlobsStorage** — Production (Azure Blob Storage)
- **CosmosDbPartitionedStorage** — Production (Azure Cosmos DB)

### Key Classes by Language

| Concept | Python | JavaScript | C# |
|---------|--------|------------|-----|
| Main class | `AgentApplication` | `AgentApplication` | `AgentApplication` |
| Context | `TurnContext` | context parameter | `ITurnContext` |
| State | `TurnState` | state parameter | `ITurnState` |
| Storage | `MemoryStorage` | `MemoryStorage` | `MemoryStorage` / `IStorage` |
