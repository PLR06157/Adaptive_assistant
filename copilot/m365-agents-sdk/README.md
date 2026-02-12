# Microsoft 365 Agents SDK — Documentation

Extracted from the [official Microsoft documentation](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/).

## Contents

| # | Document | Description |
|---|----------|-------------|
| 01 | [Overview](01-overview.md) | What is the Agents SDK, key features, supported languages, prerequisites, installation |
| 02 | [Quickstart](02-quickstart.md) | Step-by-step echo agent in Python, JavaScript, and C#/.NET with full code |
| 03 | [Agents Toolkit](03-agents-toolkit.md) | IDE integration for Visual Studio (.NET) and VS Code (JavaScript) with templates |
| 04 | [Activity Protocol](04-activity-protocol.md) | Activities, turns, TurnContext, activity types, attachments, channel-specific behavior |
| 05 | [Testing](05-testing.md) | Local testing with Microsoft 365 Agents Playground, CLI options, environment variables |
| 06 | [Deployment](06-deployment.md) | Deploy to Azure App Service, Teams manifest, Microsoft 365 integration |
| 07 | [Azure Bot Provisioning](07-azure-bot-provisioning.md) | Authentication types, Azure Bot creation, MSAL configuration |

## Quick Reference

### Supported Languages
- **C#** — .NET 8.0 SDK
- **JavaScript** — Node.js 18+
- **Python** — Python 3.9–3.11

### Key Repositories
- Main SDK: https://github.com/microsoft/Agents
- JavaScript SDK: https://github.com/microsoft/Agents-for-js

### Key Packages
- **Python:** `pip install microsoft-agents-hosting-aiohttp`
- **JavaScript:** `npm install @microsoft/agents-hosting-express`
- **C#:** Clone from GitHub (NuGet packages included in samples)

### Default Agent Endpoint
All agents listen on `http://localhost:3978/api/messages` by default.

### Testing Tools
- `agentsplayground` (standalone binary or npm `@microsoft/m365agentsplayground`)
- `teamsapptester` (npm `@microsoft/teams-app-test-tool`)

---

*Source: Microsoft Learn — last updated November 2025*
