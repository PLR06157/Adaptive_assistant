# Provisioning Azure Bot Service Manually

To run an agent on Azure, you must:
1. Create an Azure Bot
2. Configure your SDK agent for that Azure Bot
3. Deploy your SDK agent to Azure

This document covers step 1 — creating the Azure Bot. It applies to all Azure Bots regardless of the SDK language used.

---

## Authentication Types

The Azure Bot Service supports three authentication types:

### 1. Client Secret (Single Tenant)
- Traditional secret-based authentication
- Good for **local testing** (if your tenant allows secrets) with devtunnel

### 2. User Managed Identity
- Provides **better security** than client secrets
- Recommended for production workloads

### 3. Federated Credentials (Single Tenant)
- Required if creating a **Teams agent that requires SSO**
- Uses identity federation instead of secrets

---

## Which Authentication Type Should You Use?

| Scenario | Recommended Type |
|----------|-----------------|
| Teams agent requiring SSO | **Federated Credentials** |
| Better security for production | **User Managed Identity** |
| Local testing with devtunnel | **Client Secret** |

---

## Next Steps by Language

### .NET
- [Configuring your agent for your Azure Bot](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/microsoft-authentication-library-configuration-options) — MSAL configuration options

### JavaScript
- [Configuring your agent for your Azure Bot](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/azure-bot-authentication-for-javascript) — Azure Bot authentication for JS

### General
- [Deploying your agent to Azure](06-deployment.md)
- [OAuth using Federated Credentials](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/azure-bot-user-authorization-federated-credentials)
- [Configure your .NET agent for OAuth](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/agent-oauth-configuration-dotnet)
