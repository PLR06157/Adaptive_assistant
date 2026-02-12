# Microsoft 365 Agents Toolkit

The Agents Toolkit provides IDE-integrated templates and tooling for creating, testing, and deploying agents using the Microsoft 365 Agents SDK.

---

## Visual Studio (.NET)

### Prerequisites

1. **Agents Toolkit Extension** for Visual Studio (install from VS marketplace)
2. **Azure OpenAI Model** — obtain from Azure AI Foundry portal:
   - Model Name
   - Target URI
   - API Key

### Create a New Project

1. Open Visual Studio > **Create a new project**
2. Search for "agent" to find Microsoft 365 Agent Toolkit templates
3. Select **Microsoft 365 Agents** > **Next**
4. Name your agent, set location and solution name > **Create**

### Select Agent Type

Available templates:
- **Weather Agent** — Prebuilt sample with Semantic Kernel orchestration and Azure OpenAI integration
- **Empty Agent** — Start without a model or orchestrator

For this guide, select **Weather Agent** > **Next**.

### Configure LLM Service

1. Select **Azure OpenAI** as the LLM service
2. Enter:
   - **Azure OpenAI service key**
   - **Azure OpenAI endpoint**
   - **Azure OpenAI deployment name**
3. Select **Create** to generate the project

The toolkit creates a fully configured project with all necessary code scaffolding.

### Testing Options

#### Option 1: Local Testing with Agents Playground
1. Set debug target to **Microsoft 365 Agents Playground**
2. Playground opens in browser at localhost
3. Send messages to test agent behavior

#### Option 2: Debug in Microsoft Teams / Microsoft 365 Copilot
1. Select a Teams/Copilot debugging target
2. Wait for the switch to Microsoft Teams
3. You'll be prompted to add your agent in the Teams Client
4. Select **Add** > **Open** to interact with the agent
5. Set breakpoints as needed for debugging

---

## Visual Studio Code (JavaScript)

### Prerequisites

1. **Agents Toolkit Extension** for VS Code (install from marketplace)
2. **Azure OpenAI Model** from Azure AI Foundry portal:
   - Model Name
   - Target URI
   - Key

### Create a New Project

1. Open the Microsoft 365 Agents Toolkit extension side panel in VS Code
2. Select **Create a New Agent/App**
3. Choose to start from a template or from samples
4. Select **Custom Engine Agent**

### Available Templates

- **Basic Custom Engine Agent** — No prebuilt components; requires adding an AI orchestrator (Semantic Kernel or LangChain) and knowledge
- **Weather Agent** — Uses LangChain and Azure AI Foundry (recommended for getting started)

### Configure Weather Agent

1. Select **Weather Agent** template
2. Select **Azure OpenAI** as LLM service
3. Enter:
   - Azure OpenAI service **Key**
   - **Target URI**
   - **Model Name** (found under **My assets** > **Models and endpoints** in Foundry portal)
4. Choose **JavaScript** or **TypeScript**
5. Select project folder location
6. Enter Application Name
7. Authenticate — sign in with Microsoft 365 logo on toolbar

### Debug in Agents Playground

1. Select **Debug in Microsoft 365 Agents Playground**
2. Wait while local machine prepares required components
3. Playground opens in browser
4. Test with queries like: "What is the weather in [location] tomorrow?"
5. Agent responds with adaptive cards containing weather info

### Debug in Microsoft 365 Copilot

1. Switch debug target to Copilot
2. Press **F5** or select **Debug**
3. Wait while toolkit automatically:
   - Creates app registration in Azure AD
   - Configures Azure Bot Service record
   - Deploys project to your tenant with manifest
4. Microsoft 365 Copilot loads
5. Ask questions directly in Copilot
6. Full debugging with breakpoints available

**Prerequisite:** Access to a tenant with Microsoft 365 Copilot enabled.

---

## Summary

After completing the toolkit walkthrough, you will have:
- Created a new Microsoft 365 Agents project using the Agents Toolkit
- Tested the agent locally using the Microsoft 365 Agents Playground
- Deployed the agent for debugging in Teams or Microsoft 365 Copilot

## Technical Details

| Component | Details |
|-----------|---------|
| **Framework (.NET)** | .NET 8.0 |
| **Orchestration** | Semantic Kernel (.NET), LangChain (JS) |
| **LLM Options** | Azure OpenAI, Azure AI Foundry |
| **IDEs** | Visual Studio (.NET), VS Code (JS/TS) |
| **Templates** | Weather Agent, Empty Agent / Basic Custom Engine Agent |
| **Testing** | Local Playground, Microsoft Teams, Microsoft 365 Copilot |
