"""
Error message definitions for SharePoint Search MCP Server
Provides natural language error messages that are easy for AI agents to understand
"""

from enum import Enum


class ErrorCategory(Enum):
    """Error category definitions"""

    AUTHENTICATION = "authentication"
    AUTHORIZATION = "authorization"
    NETWORK = "network"
    SEARCH_QUERY = "search_query"
    FILE_NOT_FOUND = "file_not_found"
    CONFIGURATION = "configuration"
    UNKNOWN = "unknown"


class SharePointError(Exception):
    """Custom exception class for SharePoint operations"""

    def __init__(
        self,
        category: ErrorCategory,
        message: str,
        solution: str,
        original_error: Exception | None = None,
    ):
        self.category = category
        self.message = message
        self.solution = solution
        self.original_error = original_error
        super().__init__(self.get_formatted_message())

    def get_formatted_message(self) -> str:
        """Get formatted error message for AI agents"""
        return f"{self.message} {self.solution}"


def get_authentication_error(original_error: Exception) -> SharePointError:
    """Generate authentication error message"""
    error_str = str(original_error).lower()

    if "certificate" in error_str or "private_key" in error_str:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="Failed to load certificate or private key.",
            solution="Please verify that the certificate file path and private key file path, or their contents, are correctly configured.",
            original_error=original_error,
        )
    elif "401" in error_str or "unauthorized" in error_str:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="SharePoint authentication failed.",
            solution="Please verify the tenant ID, client ID, and certificate settings, and contact your administrator about the app registration status.",
            original_error=original_error,
        )
    else:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="An error occurred during authentication.",
            solution="Please check your configuration or contact your administrator.",
            original_error=original_error,
        )


def get_authorization_error(original_error: Exception) -> SharePointError:
    """Generate authorization error message"""
    return SharePointError(
        category=ErrorCategory.AUTHORIZATION,
        message="Access to SharePoint was denied.",
        solution="Please check the app permissions or request SharePoint site access from your administrator.",
        original_error=original_error,
    )


def get_network_error(original_error: Exception) -> SharePointError:
    """Generate network error message"""
    error_str = str(original_error).lower()

    if "timeout" in error_str:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="Connection to SharePoint server timed out.",
            solution="Please check your network connection or try again after a few moments.",
            original_error=original_error,
        )
    elif "connection" in error_str:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="Could not connect to SharePoint server.",
            solution="Please verify your internet connection and site URL.",
            original_error=original_error,
        )
    else:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="A network communication error occurred.",
            solution="Please check your network connection and try again.",
            original_error=original_error,
        )


def get_search_query_error(original_error: Exception) -> SharePointError:
    """Generate search query error message"""
    return SharePointError(
        category=ErrorCategory.SEARCH_QUERY,
        message="An error occurred while processing the search query.",
        solution="Please try different search keywords or specify more specific search criteria.",
        original_error=original_error,
    )


def get_file_not_found_error(
    file_path: str | None, original_error: Exception
) -> SharePointError:
    """Generate file not found error message"""
    if file_path:
        message = f"The specified file was not found: {file_path}"
    else:
        message = "The requested file was not found."

    return SharePointError(
        category=ErrorCategory.FILE_NOT_FOUND,
        message=message,
        solution="Please verify the file path is correct or obtain the correct path from the latest search results.",
        original_error=original_error,
    )


def get_configuration_error(original_error: Exception) -> SharePointError:
    """Generate configuration error message"""
    return SharePointError(
        category=ErrorCategory.CONFIGURATION,
        message="There is a problem with the SharePoint configuration.",
        solution="Please check the environment variable settings and ensure all required configuration items are correctly set.",
        original_error=original_error,
    )


def get_unknown_error(original_error: Exception) -> SharePointError:
    """Generate unknown error message"""
    return SharePointError(
        category=ErrorCategory.UNKNOWN,
        message="An unexpected error occurred.",
        solution="Please check your configuration or contact your administrator.",
        original_error=original_error,
    )


def handle_sharepoint_error(error: Exception, context: str = "") -> SharePointError:
    """
    Classify SharePoint-related errors into appropriate categories and generate natural language messages

    Args:
        error: The exception that occurred
        context: The context where the error occurred ("auth", "search", "download", etc.)

    Returns:
        SharePointError: Naturalized error message
    """
    error_str = str(error).lower()

    # Classification by HTTP status code
    if hasattr(error, "response") and hasattr(error.response, "status_code"):
        status_code = error.response.status_code
        if status_code == 401:
            return get_authentication_error(error)
        elif status_code == 403:
            return get_authorization_error(error)
        elif status_code == 404 and context == "download":
            return get_file_not_found_error(None, error)

    # Classification by error message content
    if any(
        keyword in error_str
        for keyword in ["certificate", "private_key", "jwt", "token", "auth"]
    ):
        return get_authentication_error(error)
    elif any(
        keyword in error_str
        for keyword in ["403", "forbidden", "access denied", "permission"]
    ):
        return get_authorization_error(error)
    elif any(
        keyword in error_str for keyword in ["timeout", "connection", "network", "dns"]
    ):
        return get_network_error(error)
    elif (
        any(keyword in error_str for keyword in ["query", "search"])
        and context == "search"
    ):
        return get_search_query_error(error)
    elif (
        any(keyword in error_str for keyword in ["not found", "404"])
        and context == "download"
    ):
        return get_file_not_found_error(None, error)
    elif any(
        keyword in error_str
        for keyword in ["config", "validation", "missing", "required"]
    ):
        return get_configuration_error(error)
    else:
        return get_unknown_error(error)
