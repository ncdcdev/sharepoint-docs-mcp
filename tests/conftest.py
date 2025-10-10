import pytest
from unittest.mock import Mock, patch
import os
from src.config import SharePointConfig


@pytest.fixture
def mock_config():
    """Mock SharePoint configuration for testing"""
    config = Mock(spec=SharePointConfig)
    config.site_url = "https://test.sharepoint.com/sites/test"
    config.tenant_id = "test-tenant-id"
    config.client_id = "test-client-id"
    config.certificate_path = None
    config.certificate_text = "mock-certificate"
    config.private_key_path = None
    config.private_key_text = "mock-private-key"
    config.default_max_results = 20
    config.allowed_file_extensions = ["pdf", "docx", "xlsx"]
    config.search_tool_description = "Test search tool"
    config.download_tool_description = "Test download tool"

    # OAuth認証設定
    config.auth_mode = "certificate"
    config.is_oauth_mode = False
    config.is_certificate_mode = True
    config.oauth_client_id = ""
    config.oauth_redirect_uri = "http://localhost:8000/oauth/callback"
    config.token_cache_path = ".sharepoint_tokens.json"

    # Mock validation method
    config.validate.return_value = []

    return config


@pytest.fixture
def mock_env_vars():
    """Mock environment variables for testing (certificate mode)"""
    env_vars = {
        "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
        "SHAREPOINT_SITE_NAME": "test",
        "SHAREPOINT_TENANT_ID": "test-tenant-id",
        "SHAREPOINT_CLIENT_ID": "test-client-id",
        "SHAREPOINT_AUTH_MODE": "certificate",
        "SHAREPOINT_CERTIFICATE_TEXT": "mock-certificate",
        "SHAREPOINT_PRIVATE_KEY_TEXT": "mock-private-key",
    }

    with patch.dict(os.environ, env_vars, clear=True):
        yield env_vars


@pytest.fixture
def mock_sharepoint_auth():
    """Mock SharePoint authentication"""
    with patch("src.sharepoint_auth.SharePointCertificateAuth") as mock_auth:
        auth_instance = Mock()
        auth_instance.get_access_token.return_value = "mock-access-token"
        mock_auth.return_value = auth_instance
        yield auth_instance


@pytest.fixture
def mock_sharepoint_client():
    """Mock SharePoint search client"""
    with patch("src.sharepoint_search.SharePointSearchClient") as mock_client:
        client_instance = Mock()
        client_instance.search_documents.return_value = [
            {
                "title": "Test Document 1",
                "path": "/sites/test/documents/test1.pdf",
                "size": 1024,
                "modified": "2024-01-01T00:00:00Z",
                "extension": "pdf",
                "summary": "Test document summary",
            }
        ]
        client_instance.download_file.return_value = b"mock file content"
        mock_client.return_value = client_instance
        yield client_instance