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

# Don't include offline_access - MSAL adds it automatically!
SCOPES = [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/User.Read",
]

def get_refresh_token():
    """Run this ONCE to get a refresh token"""
    
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY
    )
    
    print("Starting device code authentication...")
    print("You'll need to login with your M365 Copilot account\n")
    
    # Device code flow automatically requests offline_access
    flow = app.initiate_device_flow(scopes=SCOPES)
    
    if "user_code" not in flow:
        print(f"Failed to create device flow: {flow}")
        return False
    
    print(flow["message"])
    print("\nWaiting for authentication...")
    
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        print("\n✓ Authentication successful!")
        
        # Check if we got a refresh token
        if "refresh_token" not in result:
            print("⚠️  Warning: No refresh token received!")
            print("This might happen if your Azure app doesn't allow it.")
            print("\nTo fix:")
            print("1. Go to Azure Portal → Your App Registration")
            print("2. Click 'Authentication'")
            print("3. Under 'Advanced settings', set 'Allow public client flows' to YES")
            return False
        
        # Save the credentials
        credentials = {
            "refresh_token": result["refresh_token"],
            "access_token": result["access_token"],
            "scope": result.get("scope", "").split(" "),
            "token_type": result.get("token_type", "Bearer")
        }
        
        with open("copilot_credentials.json", "w") as f:
            json.dump(credentials, f, indent=2)
        
        print("✓ Credentials saved to copilot_credentials.json")
        print(f"✓ Refresh token obtained (expires in ~90 days)")
        print("\n⚠️  SECURITY WARNING:")
        print("   - Keep copilot_credentials.json SECRET")
        print("   - Don't commit it to git")
        print("   - Store it securely in production (e.g., Azure Key Vault)")
        print("\n✓ You can now run copilot_backend.py without browser!")
        return True
    else:
        print(f"\n✗ Authentication failed")
        print(f"Error: {result.get('error')}")
        print(f"Description: {result.get('error_description')}")
        
        if "AADSTS" in str(result.get('error_description', '')):
            print("\nCommon issues:")
            print("- Check that your Azure app has 'Allow public client flows' enabled")
            print("- Verify the CLIENT_ID and TENANT_ID are correct")
            print("- Make sure you're logging in with the right M365 account")
        
        return False

if __name__ == "__main__":
    print("=" * 70)
    print("Microsoft 365 Copilot - One-Time Authentication Setup")
    print("=" * 70)
    print()
    
    success = get_refresh_token()
    
    if success:
        print("\n" + "=" * 70)
        print("✓ Setup complete! Next steps:")
        print("=" * 70)
        print("1. Run: python3 copilot_backend.py")
        print("2. The backend will work without any browser prompts")
        print("3. Refresh token auto-renews for 90+ days")
        print("=" * 70)
    
    exit(0 if success else 1)
