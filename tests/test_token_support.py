"""Tests for direct token support in Authorization header"""

import pytest
from unittest.mock import Mock, patch
from fastmcp import Context
from fastmcp.server.auth import AccessToken

from src.server import _get_token_from_request


class TestGetTokenFromRequest:
    """_get_token_from_request function tests"""

    @pytest.mark.unit
    def test_get_token_from_authorization_header(self):
        """Test token retrieval from Authorization header"""
        # Mock Context with HTTP request
        mock_ctx = Mock(spec=Context)
        mock_request = Mock()
        mock_request.headers = {"Authorization": "Bearer test-token-from-header"}
        mock_ctx.get_http_request.return_value = mock_request

        token = _get_token_from_request(mock_ctx)

        assert token == "test-token-from-header"
        mock_ctx.get_http_request.assert_called_once()

    @pytest.mark.unit
    def test_get_token_from_authorization_header_without_bearer_prefix(self):
        """Test that token without 'Bearer ' prefix is ignored"""
        # Mock Context with HTTP request
        mock_ctx = Mock(spec=Context)
        mock_request = Mock()
        mock_request.headers = {"Authorization": "test-token-without-prefix"}
        mock_ctx.get_http_request.return_value = mock_request

        with patch("src.server.get_access_token", return_value=None):
            token = _get_token_from_request(mock_ctx)

        # Should fallback to None since no valid header and no OAuth token
        assert token is None

    @pytest.mark.unit
    def test_get_token_from_fastmcp_oauth_context(self):
        """Test token retrieval from FastMCP OAuth context"""
        mock_ctx = Mock(spec=Context)
        mock_ctx.get_http_request.side_effect = RuntimeError("Not in HTTP context")

        mock_access_token = Mock(spec=AccessToken)
        mock_access_token.token = "test-oauth-token"

        with patch("src.server.get_access_token", return_value=mock_access_token):
            token = _get_token_from_request(mock_ctx)

        assert token == "test-oauth-token"

    @pytest.mark.unit
    def test_priority_authorization_header_over_oauth(self):
        """Test that Authorization header takes priority over OAuth context"""
        # Mock Context with both header and OAuth token available
        mock_ctx = Mock(spec=Context)
        mock_request = Mock()
        mock_request.headers = {"Authorization": "Bearer header-token"}
        mock_ctx.get_http_request.return_value = mock_request

        mock_access_token = Mock(spec=AccessToken)
        mock_access_token.token = "oauth-token"

        with patch("src.server.get_access_token", return_value=mock_access_token):
            token = _get_token_from_request(mock_ctx)

        # Authorization header should take priority
        assert token == "header-token"

    @pytest.mark.unit
    def test_no_token_available(self):
        """Test when no token is available from any source"""
        mock_ctx = Mock(spec=Context)
        mock_ctx.get_http_request.side_effect = RuntimeError("Not in HTTP context")

        with patch("src.server.get_access_token", return_value=None):
            token = _get_token_from_request(mock_ctx)

        assert token is None

    @pytest.mark.unit
    def test_no_context_provided(self):
        """Test when no context is provided (ctx=None)"""
        mock_access_token = Mock(spec=AccessToken)
        mock_access_token.token = "oauth-token"

        with patch("src.server.get_access_token", return_value=mock_access_token):
            token = _get_token_from_request(None)

        assert token == "oauth-token"

    @pytest.mark.unit
    def test_attribute_error_handling(self):
        """Test handling of AttributeError when accessing HTTP context"""
        mock_ctx = Mock(spec=Context)
        mock_ctx.get_http_request.side_effect = AttributeError("No such attribute")

        mock_access_token = Mock(spec=AccessToken)
        mock_access_token.token = "oauth-token"

        with patch("src.server.get_access_token", return_value=mock_access_token):
            token = _get_token_from_request(mock_ctx)

        assert token == "oauth-token"

    @pytest.mark.unit
    def test_empty_authorization_header(self):
        """Test empty Authorization header"""
        mock_ctx = Mock(spec=Context)
        mock_request = Mock()
        mock_request.headers = {"Authorization": ""}
        mock_ctx.get_http_request.return_value = mock_request

        with patch("src.server.get_access_token", return_value=None):
            token = _get_token_from_request(mock_ctx)

        assert token is None

    @pytest.mark.unit
    def test_bearer_prefix_case_sensitive(self):
        """Test that 'Bearer ' prefix is case-sensitive"""
        mock_ctx = Mock(spec=Context)
        mock_request = Mock()
        mock_request.headers = {"Authorization": "bearer lowercase-token"}
        mock_ctx.get_http_request.return_value = mock_request

        with patch("src.server.get_access_token", return_value=None):
            token = _get_token_from_request(mock_ctx)

        # Should not match lowercase "bearer"
        assert token is None
