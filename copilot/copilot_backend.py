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

# Same specific scopes (without offline_access for token refresh)
SCOPES = [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/User.Read",
]

CREDENTIALS_FILE = "copilot_credentials.json"

def load_credentials():
    """Load saved credentials"""
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"\n❌ {CREDENTIALS_FILE} not found!\n"
            f"Run 'python3 get_refresh_token.py' first to authenticate."
        )
    
    with open(CREDENTIALS_FILE, "r") as f:
        return json.load(f)

def save_credentials(credentials):
    """Save updated credentials"""
    with open(CREDENTIALS_FILE, "w") as f:
        json.dump(credentials, f, indent=2)

def get_access_token():
    """Get access token using refresh token (NO BROWSER NEEDED)"""
    
    print("Loading saved credentials...")
    creds = load_credentials()
    refresh_token = creds.get("refresh_token")
    
    if not refresh_token:
        raise ValueError("No refresh token found. Re-run get_refresh_token.py")
    
    print("Getting fresh access token...")
    
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY
    )
    
    # Use refresh token to get new access token
    result = app.acquire_token_by_refresh_token(
        refresh_token=refresh_token,
        scopes=SCOPES
    )
    
    if "access_token" in result:
        print("✓ Access token obtained!")
        
        # Update credentials file with new tokens
        creds["access_token"] = result["access_token"]
        
        # Microsoft may issue a new refresh token
        if "refresh_token" in result:
            creds["refresh_token"] = result["refresh_token"]
            print("✓ Refresh token renewed")
        
        save_credentials(creds)
        return result["access_token"]
    else:
        error_msg = result.get('error_description', result.get('error'))
        raise Exception(f"Token refresh failed: {error_msg}")

def create_conversation(token):
    """Create a new Copilot conversation"""
    url = "https://graph.microsoft.com/beta/copilot/conversations"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print(f"\nCreating conversation...")
    response = requests.post(url, headers=headers, json={})
    
    print(f"Status: {response.status_code}")
    
    if response.status_code == 201:
        print("✓ Conversation created")
        return response.json()
    else:
        raise Exception(
            f"Failed to create conversation: {response.status_code}\n"
            f"Response: {response.text}"
        )

def send_message(token, conversation_id, message):
    """Send a message to Copilot"""
    url = f"https://graph.microsoft.com/beta/copilot/conversations/{conversation_id}/chat"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    body = {"message": {"text": message}}
    
    print(f"\nSending message to Copilot...")
    response = requests.post(url, headers=headers, json=body)
    
    print(f"Status: {response.status_code}")
    
    if response.status_code == 200:
        print("✓ Response received")
        return response.json()
    else:
        raise Exception(
            f"Failed to send message: {response.status_code}\n"
            f"Response: {response.text}"
        )

def chat_with_copilot(message):
    """
    Main function - call this from your backend
    
    Args:
        message: Your question/prompt for Copilot
    
    Returns:
        dict: Copilot's response
    """
    
    token = get_access_token()
    conversation = create_conversation(token)
    conversation_id = conversation['id']
    response = send_message(token, conversation_id, message)
    
    return response

if __name__ == "__main__":
    print("=" * 70)
    print("Microsoft 365 Copilot Chat API - Backend Test")
    print("=" * 70)
    
    try:
        # Test message
        test_message = "What are the top 3 AI trends in 2025?"
        print(f"\nTest message: '{test_message}'")
        
        result = chat_with_copilot(test_message)
        
        print("\n" + "=" * 70)
        print("COPILOT RESPONSE:")
        print("=" * 70)
        print(json.dumps(result, indent=2))
        print("=" * 70)
        print("\n✓ Success! Backend authentication working!")
        
    except FileNotFoundError as e:
        print(f"\n{e}")
        print("\nRun this first:")
        print("  python3 get_refresh_token.py")
        
    except Exception as e:
        print(f"\n✗ Error: {e}")