# Microsoft 365 Agents SDK - Overview

## What is the Microsoft 365 Agents SDK?

With the Microsoft 365 Agents SDK, you can create agents deployable to channels of your choice, such as Microsoft 365 Copilot, Microsoft Teams, Web & Custom Apps and more, with scaffolding to handle the required communication. Developers can use the AI Services of their choice, and make the agents they build available using the channel management capabilities of the SDK.

## Key Features

Developers need the flexibility to integrate agents from any provider or technology stack into their enterprise systems. The Agents SDK simplifies the implementation of agentic patterns using the AI of their choice, allowing them to select one or more services, models, or agents to meet their specific requirements.

**Use the Agents SDK to:**

1. **Quickly build an agent container** with state, storage, and the ability to manage activities and events. Deploy this container across any channel, such as Microsoft 365 Copilot or Microsoft Teams.
2. **Implement agentic patterns** without being restricted to a specific technology stack. The Agents SDK is agnostic about the AI you choose.
3. **Customize your agent** to align with the specific behaviors of clients, such as Microsoft Teams.

## Supported Languages

| Language | Requirements |
|----------|-------------|
| **C#** | .NET 8.0 SDK |
| **JavaScript** | Node.js version 18+ |
| **Python** | Python 3.9 to 3.11 |

## Important Terms

- **Turn** — A unit of work done by the agent. It can be a single message or a series of messages. Developers work with turns and manage the data between them.
- **Activity** — One of a number of interaction types managed by the agent.
- **Messages** — A message is one type of activity that can be sent to the agent. It can be a single message or a series of messages.

## Create an Agent (C# Example)

```csharp
builder.AddAgent(sp =>
{
    var agent = new AgentApplication(sp.GetRequiredService<AgentApplicationOptions>());
    agent.OnActivity(ActivityTypes.Message, async (turnContext, turnState, cancellationToken) =>
    {
        var text = turnContext.Activity.Text;
        await turnContext.SendActivityAsync(MessageFactory.Text($"Echo: {text}"), cancellationToken);
    });
});
```

This creates a new agent, listens for a message type activity and sends a message back. From here, you can add your chosen custom AI Services (e.g., Azure Foundry or OpenAI Agents) and Orchestration (e.g., Semantic Kernel).

## Prerequisites by Language

### C#
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download)
- Knowledge of ASP.NET Core
- Knowledge of asynchronous programming in C#

### JavaScript
- [Node.js](https://nodejs.org/) v18+
- Knowledge of asynchronous programming in JavaScript
- [Visual Studio Code](https://www.visualstudio.com/downloads) (optional)

### Python
- Python version 3.9+

## Download and Install

### C#
Clone the [Agents GitHub repo](https://github.com/Microsoft/Agents) locally. The repo contains SDK source libraries and samples.

### JavaScript
Clone the [Agents GitHub repo](https://github.com/Microsoft/Agents/). Read more about installing the Agents SDK Node.js packages on the [Agents-for-js repo](https://github.com/microsoft/Agents-for-js).

### Python
Clone the [Agents GitHub repo](https://github.com/Microsoft/Agents/). Installing the samples installs needed packages for the SDK.

## GitHub Resources

- Main repo: https://github.com/microsoft/Agents
- JavaScript-specific: https://github.com/microsoft/Agents-for-js
