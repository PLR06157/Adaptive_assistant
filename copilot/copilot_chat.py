import msal
import requests
from dotenv import load_dotenv
from typing import Optional
import os


load_dotenv()

class ConfigurationError(RuntimeError):
    """Raised when required configuration is missing."""

def _read_env(name: str, *, required: bool = True, default: Optional[str] = None) -> str:
    value = os.getenv(name, default)
    if required and not value:
        raise ConfigurationError(
            f"Missing required configuration for {name}. "
            "Set it in your environment or .env file."
        )
    return value or ""

# Replace with YOUR values from Azure
CLIENT_ID = _read_env("COPILOT_CLIENT_ID")
CLIENT_SECRET = _read_env("COPILOT_SECRET")
TENANT_ID = _read_env("COPILOT_TENANT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# These are the scopes we need for Copilot Chat API
SCOPES = [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/User.Read",
]

def get_access_token_interactive():
    """Get access token using device code flow (you'll login in browser)"""
    
    # Create public client application
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY
    )
    
    # First, try to get token from cache
    accounts = app.get_accounts()
    if accounts:
        print("Found cached account, trying to get token silently...")
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("✓ Got token from cache!")
            return result["access_token"]
    
    # If no cache, use device code flow
    print("\nStarting device code authentication...")
    print("You will need to login with your M365 Copilot account")
    print("-" * 60)
    
    flow = app.initiate_device_flow(scopes=SCOPES)
    
    if "user_code" not in flow:
        raise ValueError(f"Failed to create device flow: {flow}")
    
    print(flow["message"])
    print("-" * 60)
    
    # Wait for user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        print("\n✓ Successfully authenticated!")
        return result["access_token"]
    else:
        print("\n✗ Authentication failed")
        print(f"Error: {result.get('error')}")
        print(f"Description: {result.get('error_description')}")
        return None

def create_conversation(token):
    """Create a new Copilot conversation"""
    url = "https://graph.microsoft.com/beta/copilot/conversations"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print(f"\nPOST {url}")
    response = requests.post(url, headers=headers, json={})
    
    print(f"Status Code: {response.status_code}")
    
    if response.status_code == 201:
        print("✓ Conversation created successfully!")
        return response.json()
    else:
        print("✗ Failed to create conversation")
        print(f"Response: {response.text}")
        return None

def send_message(token, conversation_id, message):
    """Send a message to Copilot"""
    url = f"https://graph.microsoft.com/beta/copilot/conversations/{conversation_id}/chat"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    body = {
        "message": {
            "text": message
        }
    }
    
    print(f"\nPOST {url}")
    print(f"Message: '{message}'")
    
    response = requests.post(url, headers=headers, json=body)
    
    print(f"Status Code: {response.status_code}")
    
    if response.status_code == 200:
        print("✓ Message sent successfully!")
        return response.json()
    else:
        print("✗ Failed to send message")
        print(f"Response: {response.text}")
        return None

if __name__ == "__main__":
    print("=" * 70)
    print("Microsoft 365 Copilot Chat API Test (User Authentication)")
    print("=" * 70)
    
    # Step 1: Authenticate as user
    print("\n[Step 1] Authenticating with your M365 account...")
    token = get_access_token_interactive()
    
    if not token:
        print("\n❌ Authentication failed")
        exit(1)
    
    # Step 2: Create conversation
    print("\n[Step 2] Creating Copilot conversation...")
    conversation = create_conversation(token)
    
    if not conversation:
        print("\n❌ Failed to create conversation")
        print("\nPossible reasons:")
        print("1. You need Private Preview access to Chat API")
        print("2. Your account needs admin consent for the permissions")
        print("3. Copilot license might not be properly assigned")
        exit(1)
    
    conversation_id = conversation.get('id')
    print(f"✓ Conversation ID: {conversation_id}")
    
    # Step 3: Send message
    print("\n[Step 3] Sending message to Copilot...")
    message = "What are the top 3 AI trends in 2025?"
    response = send_message(token, conversation_id, message)
    
    if response:
        print("\n" + "=" * 70)
        print("COPILOT RESPONSE")
        print("=" * 70)
        print(json.dumps(response, indent=2))
        print("=" * 70)
    else:
        print("\n❌ No response received")
