"""
Authentication module for Google Slides Templater.
Provides credential management for Google API access.
"""

import json
import logging
import os
import threading
import time
from pathlib import Path
from typing import List, Optional, Union

import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as OAuth2Credentials
from google.oauth2.service_account import \
    Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

DEFAULT_SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]


class AuthConfig(BaseModel):
    """Configuration for authentication"""

    service_account_file: Optional[str] = None
    credentials_path: Optional[str] = None
    token_path: Optional[str] = None
    scopes: List[str] = Field(default_factory=lambda: DEFAULT_SCOPES.copy())
    use_application_default: bool = True
    oauth_timeout: int = Field(default=300, ge=30, le=3600)


class SlidesAPIError(Exception):
    """Base exception for Slides API"""

    pass


class AuthenticationError(SlidesAPIError):
    """Authentication failed"""

    pass


class TokenRefreshError(SlidesAPIError):
    """Token refresh failed"""

    pass


def _is_safe_path(filepath: str, base_dir: str = ".") -> bool:
    """Check if filepath is safe (no path traversal)"""
    try:
        base_path = Path(base_dir).resolve()
        target_path = Path(filepath).resolve()
        return target_path.is_relative_to(base_path)
    except (ValueError, OSError):
        return False


class CredentialManager:
    """
    Manager for Google API credentials.
    Supports different authentication methods.
    """

    def __init__(self, scopes: Optional[List[str]] = None):
        """
        Initialize credential manager.

        Args:
            scopes: OAuth scopes for API access
        """
        self.scopes = scopes or DEFAULT_SCOPES.copy()

    def from_service_account_file(self, service_account_file: str) -> "Credentials":
        """
        Create credentials from service account file.

        Args:
            service_account_file: Path to service account JSON file

        Returns:
            Credentials object

        Raises:
            AuthenticationError: If authentication failed
        """
        if not _is_safe_path(service_account_file):
            raise AuthenticationError("Unsafe file path")

        if not os.path.exists(service_account_file):
            raise AuthenticationError(
                f"Service account file not found: {service_account_file}"
            )

        try:
            service_credentials = ServiceAccountCredentials.from_service_account_file(
                service_account_file, scopes=self.scopes
            )

            logger.info(
                f"Loaded service account credentials from {service_account_file}"
            )
            return Credentials(service_credentials, auth_method="service_account")

        except Exception as e:
            logger.error(f"Error loading service account credentials: {e}")
            raise AuthenticationError(f"Invalid service account file: {e}")

    def from_saved_token(self, token_file: str) -> Optional["Credentials"]:
        """
        Load credentials from saved token file.

        Args:
            token_file: Path to saved token file

        Returns:
            Credentials object or None if loading failed
        """
        if not _is_safe_path(token_file):
            logger.warning(f"Unsafe token file path: {token_file}")
            return None

        if not os.path.exists(token_file):
            logger.debug(f"Token file not found: {token_file}")
            return None

        try:
            with open(token_file, "r", encoding="utf-8") as f:
                token_data = json.load(f)

            required_fields = ["token", "refresh_token", "client_id", "client_secret"]
            if not all(field in token_data for field in required_fields):
                logger.warning(f"Token file missing required fields: {token_file}")
                return None

            api_client = OAuth2Credentials.from_authorized_user_info(
                token_data, self.scopes
            )

            if api_client:
                if not api_client.valid:
                    if api_client.expired and api_client.refresh_token:
                        logger.info("Token expired, refreshing...")
                        try:
                            api_client.refresh(Request())
                            self._save_token(creds, token_file)
                            logger.info("Token refreshed and saved")
                        except Exception as refresh_error:
                            logger.warning(f"Failed to refresh token: {refresh_error}")
                            return None
                    else:
                        logger.warning(f"Invalid token in {token_file}")
                        return None

                if api_client.valid:
                    logger.info(f"Loaded valid token from {token_file}")
                    return Credentials(creds, auth_method="saved_token")

            return None

        except json.JSONDecodeError as e:
            logger.warning(f"Invalid JSON in token file {token_file}: {e}")
            return None
        except Exception as e:
            logger.warning(f"Failed to load token from {token_file}: {e}")
            return None

    def from_oauth_flow(
        self,
        client_secrets_file: str,
        token_save_path: Optional[str] = None,
        use_local_server: bool = True,
        timeout: int = 300,
    ) -> "Credentials":
        """
        Create credentials through OAuth flow.

        Args:
            client_secrets_file: Path to OAuth credentials JSON file
            token_save_path: Path to save token
            use_local_server: Use local server for OAuth
            timeout: Timeout for OAuth flow

        Returns:
            Credentials object

        Raises:
            AuthenticationError: If OAuth flow failed
        """
        if not _is_safe_path(client_secrets_file):
            raise AuthenticationError("Unsafe credentials file path")

        if not os.path.exists(client_secrets_file):
            raise AuthenticationError(
                f"OAuth credentials file not found: {client_secrets_file}"
            )

        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secrets_file, self.scopes
            )

            if use_local_server:
                print("Opening browser for authentication...")
                api_client = flow.run_local_server(port=0, open_browser=True)
            else:
                print("Console authentication mode")
                print("Go to the following URL in your browser:")
                auth_url, _ = flow.authorization_url(prompt="consent")
                print(f"\n{auth_url}\n")

                auth_code = input("Enter the authorization code: ").strip()
                if not auth_code:
                    raise AuthenticationError("No authorization code provided")

                flow.fetch_token(code=auth_code)
                api_client = flow.credentials

            if token_save_path:
                self._save_token(creds, token_save_path)
                logger.info(f"Token saved to {token_save_path}")

            logger.info("OAuth authentication completed successfully")
            return Credentials(creds, auth_method="oauth")

        except Exception as e:
            logger.error(f"OAuth authentication failed: {e}")
            raise AuthenticationError(f"OAuth flow failed: {e}")

    def from_application_default(self) -> Optional["Credentials"]:
        """
        Get credentials from Application Default Credentials (ADC).

        Returns:
            Credentials object or None if ADC not available
        """
        try:
            credentials, project = google.auth.default(scopes=self.scopes)
            logger.info(f"Loaded Application Default Credentials (project: {project})")
            return Credentials(credentials, auth_method="application_default")

        except Exception as e:
            logger.debug(f"Application Default Credentials not available: {e}")
            return None

    def _save_token(self, credentials: OAuth2Credentials, token_file: str):
        """Save token to file."""
        if not _is_safe_path(token_file):
            raise AuthenticationError("Unsafe token file path")

        try:
            token_path = Path(token_file)
            token_path.parent.mkdir(parents=True, exist_ok=True)

            token_data = {
                "token": credentials.token,
                "refresh_token": credentials.refresh_token,
                "token_uri": credentials.token_uri,
                "client_id": credentials.client_id,
                "client_secret": credentials.client_secret,
                "scopes": credentials.scopes,
                "expiry": (
                    credentials.expiry.isoformat() if credentials.expiry else None
                ),
            }

            with open(token_file, "w", encoding="utf-8") as f:
                json.dump(token_data, f, indent=2)

            logger.debug(f"Token saved to {token_file}")

        except Exception as e:
            logger.error(f"Failed to save token to {token_file}: {e}")
            raise AuthenticationError(f"Failed to save token: {e}")


class Credentials:
    """
    Wrapper for Google API credentials with additional functionality.
    """

    def __init__(
        self,
        credentials: Union[OAuth2Credentials, ServiceAccountCredentials],
        auth_method: str = "unknown",
    ):
        """
        Initialize credentials wrapper.

        Args:
            credentials: Google API credentials
            auth_method: Authentication method used
        """
        self.credentials = credentials
        self.auth_method = auth_method
        self._last_refresh_time = time.time()
        self._refresh_lock = threading.Lock()

    @property
    def valid(self) -> bool:
        """Check if credentials are valid."""
        if not self.credentials:
            return False
        return self.credentials.valid

    @property
    def expired(self) -> bool:
        """Check if credentials are expired."""
        if not hasattr(self.credentials, "expired"):
            return False
        return self.credentials.expired

    def refresh_if_needed(self) -> bool:
        """
        Refresh credentials if needed (thread-safe).

        Returns:
            True if credentials were refreshed
        """
        if not self.expired:
            return False

        with self._refresh_lock:
            # Double-check after acquiring lock
            if not self.expired:
                return False

            if (
                not hasattr(self.credentials, "refresh_token")
                or not self.credentials.refresh_token
            ):
                logger.warning("Cannot refresh credentials: no refresh token")
                return False

            try:
                logger.info("Refreshing expired credentials...")
                self.credentials.refresh(Request())
                self._last_refresh_time = time.time()
                logger.info("Credentials refreshed successfully")
                return True

            except Exception as e:
                logger.error(f"Failed to refresh credentials: {e}")
                raise TokenRefreshError(f"Failed to refresh credentials: {e}")

    def ensure_valid(self):
        """
        Ensure credentials are valid, refresh if needed.

        Raises:
            AuthenticationError: If credentials are invalid and cannot be refreshed
        """
        if self.valid:
            return

        if self.refresh_if_needed():
            return

        raise AuthenticationError(f"Invalid credentials (method: {self.auth_method})")

    def get_info(self) -> dict:
        """
        Get information about credentials.

        Returns:
            Dictionary with credential information
        """
        info = {
            "auth_method": self.auth_method,
            "valid": self.valid,
            "expired": self.expired,
            "last_refresh": self._last_refresh_time,
        }

        if hasattr(self.credentials, "service_account_email"):
            info["service_account_email"] = self.credentials.service_account_email

        if hasattr(self.credentials, "refresh_token"):
            info["has_refresh_token"] = bool(self.credentials.refresh_token)

        return info


def authenticate(config: Optional[AuthConfig] = None, **kwargs) -> Credentials:
    """
    Universal authentication function.

    Args:
        config: Authentication configuration
        **kwargs: Legacy parameters for backward compatibility

    Returns:
        Credentials object

    Raises:
        AuthenticationError: If no authentication method succeeded
    """
    # Handle legacy parameters
    if config is None:
        config = AuthConfig(**kwargs)

    manager = CredentialManager(config.scopes)

    # Method 1: Service Account
    if config.service_account_file:
        try:
            logger.info("Trying service account authentication...")
            return manager.from_service_account_file(config.service_account_file)
        except Exception as e:
            logger.warning(f"Service account authentication failed: {e}")

    # Method 2: Saved token
    if config.token_path:
        logger.info(f"Trying saved token from {config.token_path}...")
        credentials = manager.from_saved_token(config.token_path)
        if credentials:
            logger.info("Using saved token")
            return credentials
        else:
            logger.info("Saved token not valid, will try OAuth flow")

    # Method 3: OAuth flow
    if config.credentials_path:
        try:
            logger.info("Starting OAuth flow...")
            return manager.from_oauth_flow(
                config.credentials_path, config.token_path, timeout=config.oauth_timeout
            )
        except Exception as e:
            logger.warning(f"OAuth authentication failed: {e}")

    # Method 4: Application Default Credentials
    if config.use_application_default:
        logger.info("Trying Application Default Credentials...")
        credentials = manager.from_application_default()
        if credentials:
            return credentials

    raise AuthenticationError(
        "Authentication failed. Please provide valid credentials:\n"
        "- service_account_file: path to service account JSON\n"
        "- credentials_path: path to OAuth client secrets JSON\n"
        "- token_path: path to saved OAuth token\n"
        "- or ensure Application Default Credentials are configured"
    )


def setup_oauth_flow(
    client_secrets_file: str,
    token_save_path: str = "token.json",
    scopes: Optional[List[str]] = None,
) -> Credentials:
    """
    Set up OAuth flow for interactive authentication.

    Args:
        client_secrets_file: Path to OAuth client secrets file
        token_save_path: Where to save the obtained token
        scopes: OAuth scopes

    Returns:
        Credentials object
    """
    manager = CredentialManager(scopes)
    return manager.from_oauth_flow(client_secrets_file, token_save_path)


def validate_credentials(credentials: Credentials) -> bool:
    """
    Validate credentials.

    Args:
        credentials: Credentials object to validate

    Returns:
        True if credentials are valid
    """
    try:
        credentials.ensure_valid()
        return True
    except Exception as e:
        logger.error(f"Credentials validation failed: {e}")
        return False


def get_credentials_info(credentials: Credentials) -> dict:
    """
    Get detailed information about credentials.

    Args:
        credentials: Credentials object

    Returns:
        Dictionary with credentials information
    """
    return credentials.get_info()


def create_service_account_template(output_file: str = "service_account_template.json"):
    """
    Create service account template file.

    Args:
        output_file: Output file name
    """
    if not _is_safe_path(output_file):
        raise ValueError("Unsafe output file path")

    template = {
        "type": "service_account",
        "project_id": "your-project-id",
        "private_key_id": "your-private-key-id",
        "private_key": "-----BEGIN PRIVATE KEY-----\nYOUR_PRIVATE_KEY\n-----END PRIVATE KEY-----\n",
        "client_email": "your-service-account@your-project-id.iam.gserviceaccount.com",
        "client_id": "your-client-id",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project-id.iam.gserviceaccount.com",
    }

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(template, f, indent=2)

    print(f"Service account template created: {output_file}")
    print(
        "Please replace placeholder values with actual credentials from Google Cloud Console"
    )


def create_oauth_template(output_file: str = "oauth_credentials_template.json"):
    """
    Create OAuth credentials template file.

    Args:
        output_file: Output file name
    """
    if not _is_safe_path(output_file):
        raise ValueError("Unsafe output file path")

    template = {
        "installed": {
            "client_id": "your-client-id.apps.googleusercontent.com",
            "project_id": "your-project-id",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": "your-client-secret",
            "redirect_uris": ["http://localhost"],
        }
    }

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(template, f, indent=2)

    print(f"OAuth credentials template created: {output_file}")
    print(
        "Please replace placeholder values with actual credentials from Google Cloud Console"
    )


def check_credentials_file(file_path: str) -> dict:
    """
    Check credentials file and determine its type.

    Args:
        file_path: Path to credentials file

    Returns:
        Dictionary with file information
    """
    if not _is_safe_path(file_path):
        return {"exists": False, "error": "Unsafe file path"}

    if not os.path.exists(file_path):
        return {"exists": False, "error": "File not found"}

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if "type" in data and data["type"] == "service_account":
            cred_type = "service_account"
            required_fields = ["project_id", "private_key", "client_email"]
        elif "installed" in data or "web" in data:
            cred_type = "oauth"
            required_fields = ["client_id", "client_secret"]
        elif "token" in data and "refresh_token" in data:
            cred_type = "saved_token"
            required_fields = ["token", "refresh_token", "client_id", "client_secret"]
        else:
            return {
                "exists": True,
                "type": "unknown",
                "error": "Unknown credentials format",
            }

        missing_fields = []
        if cred_type == "service_account":
            for field in required_fields:
                if field not in data or not data[field]:
                    missing_fields.append(field)
        elif cred_type == "oauth":
            oauth_data = data.get("installed", data.get("web", {}))
            for field in required_fields:
                if field not in oauth_data or not oauth_data[field]:
                    missing_fields.append(field)
        elif cred_type == "saved_token":
            for field in required_fields:
                if field not in data or not data[field]:
                    missing_fields.append(field)

        result = {
            "exists": True,
            "type": cred_type,
            "valid": len(missing_fields) == 0,
            "missing_fields": missing_fields,
        }

        if cred_type == "service_account":
            result["project_id"] = data.get("project_id")
            result["client_email"] = data.get("client_email")
        elif cred_type == "saved_token":
            result["has_refresh_token"] = bool(data.get("refresh_token"))
            result["expires"] = data.get("expiry")

        return result

    except json.JSONDecodeError:
        return {"exists": True, "error": "Invalid JSON format"}
    except Exception as e:
        return {"exists": True, "error": str(e)}


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    print("Google Slides Templater Authentication Test")
    print("=" * 50)

    test_files = ["service_account.json", "oauth_credentials.json", "token.json"]

    for file_path in test_files:
        print(f"\nChecking {file_path}:")
        info = check_credentials_file(file_path)

        if info["exists"]:
            print(f"   Type: {info.get('type', 'unknown')}")
            print(f"   Valid: {'✓' if info.get('valid') else '✗'}")
            if info.get("missing_fields"):
                print(f"   Missing: {info['missing_fields']}")
        else:
            print(f"   File not found")
