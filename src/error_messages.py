"""
Error message definitions for SharePoint Search MCP Server
Provides natural language error messages that are easy for AI agents to understand
"""

from enum import Enum

from src.config import config


class ErrorCategory(Enum):
    """Error category definitions"""

    AUTHENTICATION = "authentication"
    AUTHORIZATION = "authorization"
    NETWORK = "network"
    SEARCH_QUERY = "search_query"
    FILE_NOT_FOUND = "file_not_found"
    CONFIGURATION = "configuration"
    EXCEL_FILE_NOT_FOUND = "excel_file_not_found"
    EXCEL_SHEET_NOT_FOUND = "excel_sheet_not_found"
    EXCEL_INVALID_RANGE = "excel_invalid_range"
    EXCEL_INVALID_FILE = "excel_invalid_file"
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

    if "oauth/login" in error_str or "no valid access token" in error_str:
        login_url = f"{config.oauth_server_base_url.rstrip('/')}/auth/login"
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="OAuth authentication required but not completed.",
            solution=f"Please visit {login_url} to authenticate with your Microsoft account. After successful authentication, tokens will be cached and you can retry this operation.",
            original_error=original_error,
        )
    elif "certificate" in error_str or "private_key" in error_str:
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
    file_path: str | None, original_error: Exception, is_onedrive_file: bool = False
) -> SharePointError:
    """Generate file not found error message"""
    if file_path:
        message = f"The specified file was not found: {file_path}"
    else:
        message = "The requested file was not found."

    if is_onedrive_file:
        message += " This appears to be a OneDrive file."

    # OneDriveファイルの場合は特別なメッセージを提供
    error_str = str(original_error).lower()
    if is_onedrive_file:
        solution = "This appears to be a OneDrive personal file. The system tried multiple download methods including GetFileByServerRelativePath and GetFileByServerRelativeUrl. Please verify the file still exists and you have access permissions."
    elif "all download methods failed" in error_str:
        solution = "Multiple download methods were attempted but all failed. This could be due to special characters in the filename, permission issues, or the file being moved. Please try searching for the file again to get an updated path."
    else:
        solution = "Please verify the file path is correct or obtain the correct path from the latest search results."

    return SharePointError(
        category=ErrorCategory.FILE_NOT_FOUND,
        message=message,
        solution=solution,
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


def get_excel_file_not_found_error(
    file_path: str, original_error: Exception
) -> SharePointError:
    """Generate Excel file not found error message"""
    return SharePointError(
        category=ErrorCategory.EXCEL_FILE_NOT_FOUND,
        message=f"The specified Excel file was not found: {file_path}",
        solution="Please verify the file path is correct and the file exists. You can search for the file using sharepoint_docs_search with file_extensions=['xlsx'] to get the correct path.",
        original_error=original_error,
    )


def get_excel_sheet_not_found_error(
    sheet_name: str, original_error: Exception
) -> SharePointError:
    """Generate Excel sheet not found error message"""
    return SharePointError(
        category=ErrorCategory.EXCEL_SHEET_NOT_FOUND,
        message=f"The specified sheet was not found: {sheet_name}",
        solution="Run sharepoint_excel without specifying 'sheet' to list available sheets (check sheets[].name in the response), then use a valid sheet name.",
        original_error=original_error,
    )


def get_excel_invalid_range_error(
    range_spec: str, original_error: Exception
) -> SharePointError:
    """Generate Excel invalid range error message"""
    return SharePointError(
        category=ErrorCategory.EXCEL_INVALID_RANGE,
        message=f"The specified cell range is invalid: {range_spec}",
        solution="Please use a valid range format like 'A1:C10' or 'A1'. Ensure the range is within the actual bounds of the Excel file.",
        original_error=original_error,
    )


def get_excel_invalid_file_error(original_error: Exception) -> SharePointError:
    """Generate Excel invalid file error message"""
    return SharePointError(
        category=ErrorCategory.EXCEL_INVALID_FILE,
        message="The file is not a valid Excel file or is corrupted.",
        solution="Please verify the file is a valid .xlsx file. Try opening it in Excel locally to check for corruption, or re-upload the file to SharePoint.",
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


def handle_sharepoint_error(
    error: Exception,
    context: str = "",
    is_onedrive_file: bool = False,
    excel_context: dict[str, str | None] | None = None,
) -> SharePointError:
    """
    Classify SharePoint-related errors into appropriate categories and generate natural language messages

    Args:
        error: The exception that occurred
        context: The context where the error occurred ("auth", "search", "download", "excel_*", etc.)
        is_onedrive_file: Whether the operation is for OneDrive file
        excel_context: Excel operation context with file_path, sheet_name, range_spec

    Returns:
        SharePointError: Natural language error message
    """
    error_str = str(error).lower()

    # Excel操作のエラー分類（openpyxlベース）
    if context.startswith("excel_") or context == "excel_parse":
        # ValueErrorはシート名が見つからない場合
        if isinstance(error, ValueError):
            if "not found" in error_str and "sheet" in error_str:
                sheet_name = excel_context.get("sheet_name") if excel_context else None
                return get_excel_sheet_not_found_error(sheet_name or "unknown", error)

        # openpyxlの無効な座標例外
        error_type_name = type(error).__name__.lower()
        if "coordinate" in error_type_name or "invalid" in error_type_name:
            range_spec = excel_context.get("range_spec") if excel_context else None
            return get_excel_invalid_range_error(range_spec or "unknown", error)

        # ファイル形式エラー（zipfile.BadZipFile, openpyxl例外など）
        if "badzip" in error_type_name or "not a valid" in error_str or "corrupt" in error_str:
            return get_excel_invalid_file_error(error)

        # HTTP 404エラー（ファイルが見つからない）
        if hasattr(error, "response") and hasattr(error.response, "status_code"):
            status_code = error.response.status_code
            if status_code == 404:
                file_path = excel_context.get("file_path") if excel_context else ""
                return get_excel_file_not_found_error(file_path or "", error)

    # Classification by HTTP status code
    if hasattr(error, "response") and hasattr(error.response, "status_code"):
        status_code = error.response.status_code
        if status_code == 401:
            return get_authentication_error(error)
        elif status_code == 403:
            return get_authorization_error(error)
        elif status_code == 404 and context == "download":
            return get_file_not_found_error(None, error, is_onedrive_file)

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
