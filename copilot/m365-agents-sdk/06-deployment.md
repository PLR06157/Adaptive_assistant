# Deploying Your Agent to Azure

This guide covers deploying a Microsoft 365 Agents SDK agent to Azure, Microsoft Teams, and Microsoft 365 Copilot.

---

## Prerequisites

- A working agent (see [Quickstart](02-quickstart.md))
- An Azure subscription
- Azure Bot resource provisioned (see [Azure Bot Provisioning](07-azure-bot-provisioning.md))

---

## Step 1: Publish Your Agent to a Web App

An SDK agent is a web application. Deploy using standard Azure methods:

- Deploy as a **ZIP package** to an Azure App Service app
- **Visual Studio publish** to an Azure App Service app or container
- Other **container deployments** supported by Azure
- **Microsoft 365 Agents Toolkit** deployment

### Important: Identity Configuration

If using Azure App Service with **Federated Credentials** or **User Managed Identity**, add that identity under:
**Settings** > **Identity**

### Update Messaging Endpoint

After deployment:

1. Your agent will have a base URL (e.g., `example.azurewebsites.net`)
2. Navigate to your **Azure Bot resource**
3. Go to **Configuration**
4. Change the **Messaging endpoint** to:
   ```
   https://{yourwebsite}/api/messages
   ```
   Replace `{yourwebsite}` with your web app's base URL.

---

## Step 2: Test in Web Chat

1. In your Azure Bot resource, select **Test in Web Chat**
2. Send messages to your agent to verify functionality

---

## Step 3: Prepare Teams and Microsoft 365 Copilot Manifest

### Create the Manifest

1. Create an empty folder in your project (e.g., `appManifest/`)
2. Copy contents from [Teams manifest files](https://github.com/microsoft/Agents/blob/main/samples/dotnet/quickstart/appManifest)
3. Edit `manifest.json`:
   - Replace all instances of `<<AAD_APP_CLIENT_ID>>` with your Azure Bot resource's **ClientId**
   - Replace `<<BOT_DOMAIN>>` with your agent base URL
4. Create `manifest.zip` containing:
   - `manifest.json`
   - `outline.png`
   - `color.png`

---

## Step 4: Deploy to Microsoft 365

1. Ensure **Microsoft Teams** channel is added to your Azure Bot resource under **Channels**
2. Navigate to **Microsoft Admin Portal (MAC)**
3. Go to **Settings** > **Integrated Apps**
4. Select **Upload Custom App**
5. Upload the `manifest.zip` file

**Result:** After a short period, your agent will appear in Microsoft Teams and Microsoft 365 Copilot.
