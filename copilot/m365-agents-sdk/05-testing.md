# Testing Your Agent Locally

Test agents locally using the Microsoft 365 Agents Playground before deploying to Azure.

---

## Three Ways to Create an Agent

1. **Agents Toolkit** — C#, JS, or Python using Visual Studio or VS Code (includes built-in Playground support)
2. **Clone from a sample** — Clone from GitHub and open in your IDE
3. **CLI** — Create using command-line tools

> **Recommendation:** Wherever possible, start with the Microsoft 365 Agents Toolkit. It makes getting started, testing locally, and deploying easier and quicker by abstracting the manual setup of Azure Bot Service and Azure App Registrations.

---

## Installing the Agents Playground

### Option 1: Standalone Binary

**Windows:**
```bash
winget install agentsplayground
```

**Linux:**
```bash
curl -s https://raw.githubusercontent.com/OfficeDev/microsoft-365-agents-toolkit/dev/.github/scripts/install-agentsplayground-linux.sh | bash
```

### Option 2: NPM

**Global install (recommended):**
```bash
npm install -g @microsoft/m365agentsplayground
```

**Project-specific install:**
```bash
npm install -D @microsoft/m365agentsplayground
```

---

## Testing Step-by-Step

### Step 1: Create or Clone Your Agent
Create a quickstart agent or clone a sample from the repo.

### Step 2: Configure Authentication (if needed)
- **Anonymous mode** — No configuration required
- **Authenticated mode** — Configure Microsoft Entra ID app registrations for both the Agents Playground and your application

### Step 3: Configure Ports
Select an available port for your agent:
- Default: `3978`
- Alternative: Any available port

### Step 4: Run Your Agent Code
Execute your agent application in a terminal.

### Step 5: Open Agents Playground

**Basic (anonymous):**
```bash
agentsplayground -e "http://localhost:<port>/api/messages" -c "emulator"
```

**With authentication:**
```bash
agentsplayground -e "http://localhost:<port>/api/messages" -c "emulator" \
  --client-id "your-client-id" \
  --client-secret "your-client-secret" \
  --tenant-id "your-tenant-id"
```

---

## CLI Options Reference

| Option | Short | Description | Example |
|--------|-------|-------------|---------|
| `--app-endpoint` | `-e` | Agent's endpoint URL | `http://localhost:3978/api/messages` |
| `--channel-id` | `-c` | Channel type | `emulator`, `webchat`, `msteams` |
| `--client-id` | | Client ID for auth | `your-client-id` |
| `--client-secret` | | Client secret for auth | `your-client-secret` |
| `--tenant-id` | | Tenant ID for auth | `your-tenant-id` |

View all options:
```bash
agentsplayground --help
```

---

## Environment Variables Alternative

Instead of CLI options, use environment variables (CLI options take priority):

```bash
export BOT_ENDPOINT="http://localhost:<port>/api/messages"
export DEFAULT_CHANNEL_ID="emulator"
export AUTH_CLIENT_ID="your-client-id"
export AUTH_CLIENT_SECRET="your-client-secret"
export AUTH_TENANT_ID="your-tenant-id"
```

---

## Legacy Test Tool (teams-app-test-tool)

Also available for testing:

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

Expected playground output:
```
Listening on 56150
Microsoft 365 Agents Playground is being launched for you to debug the app: http://localhost:56150
started web socket client
Waiting for connection of endpoint: http://127.0.0.1:3978/api/messages
```
