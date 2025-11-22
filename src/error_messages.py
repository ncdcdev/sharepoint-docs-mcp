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
    FOLDER_NOT_FOUND = "folder_not_found"
    FILE_ALREADY_EXISTS = "file_already_exists"
    UPLOAD_FAILED = "upload_failed"
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


def get_folder_not_found_error(
    folder_path: str | None, original_error: Exception, is_onedrive: bool = False
) -> SharePointError:
    """Generate folder not found error message"""
    if folder_path:
        message = f"The specified folder was not found: {folder_path}"
    else:
        message = "The requested folder was not found."

    if is_onedrive:
        solution = (
            "Please verify the OneDrive folder path is correct. "
            "Ensure the user email and folder path are properly formatted "
            "(e.g., '@onedrive:user@domain.com:/Documents/Folder')."
        )
    else:
        solution = (
            "Please verify the folder path is correct. "
            "Use the format 'SiteName:/Folder/Path' or provide a full URL. "
            "You may need to create the folder first."
        )

    return SharePointError(
        category=ErrorCategory.FOLDER_NOT_FOUND,
        message=message,
        solution=solution,
        original_error=original_error,
    )


def get_file_already_exists_error(
    file_name: str | None, original_error: Exception
) -> SharePointError:
    """Generate file already exists error message"""
    if file_name:
        message = f"A file with the name '{file_name}' already exists in the destination folder."
    else:
        message = "A file with the same name already exists in the destination folder."

    return SharePointError(
        category=ErrorCategory.FILE_ALREADY_EXISTS,
        message=message,
        solution="Use the 'overwrite' parameter set to true to replace the existing file, or choose a different file name.",
        original_error=original_error,
    )


def get_upload_error(
    original_error: Exception, is_onedrive: bool = False
) -> SharePointError:
    """Generate upload error message"""
    error_str = str(original_error).lower()

    if "file size" in error_str or "too large" in error_str or "413" in error_str:
        return SharePointError(
            category=ErrorCategory.UPLOAD_FAILED,
            message="The file is too large to upload.",
            solution="SharePoint has a file size limit (typically 250MB for REST API). Please try uploading a smaller file or contact your administrator about size limits.",
            original_error=original_error,
        )

    location = "OneDrive" if is_onedrive else "SharePoint"
    return SharePointError(
        category=ErrorCategory.UPLOAD_FAILED,
        message=f"Failed to upload the file to {location}.",
        solution="Please verify you have write permissions to the destination folder and the folder path is correct.",
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
    error: Exception, context: str = "", is_onedrive_file: bool = False
) -> SharePointError:
    """
    Classify SharePoint-related errors into appropriate categories and generate natural language messages

    Args:
        error: The exception that occurred
        context: The context where the error occurred ("auth", "search", "download", etc.)

    Returns:
        SharePointError: Natural language error message
    """
    error_str = str(error).lower()

    # Classification by HTTP status code
    if hasattr(error, "response") and hasattr(error.response, "status_code"):
        status_code = error.response.status_code
        if status_code == 401:
            return get_authentication_error(error)
        elif status_code == 403:
            return get_authorization_error(error)
        elif status_code == 404:
            if context == "download":
                return get_file_not_found_error(None, error, is_onedrive_file)
            elif context == "upload":
                return get_folder_not_found_error(None, error, is_onedrive_file)
        elif status_code == 409 and context == "upload":
            return get_file_already_exists_error(None, error)
        elif status_code == 413 and context == "upload":
            return get_upload_error(error, is_onedrive_file)

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
        return get_file_not_found_error(None, error, is_onedrive_file)
    elif (
        any(keyword in error_str for keyword in ["not found", "404"])
        and context == "upload"
    ):
        return get_folder_not_found_error(None, error, is_onedrive_file)
    elif ("already exists" in error_str or "409" in error_str) and context == "upload":
        return get_file_already_exists_error(None, error)
    elif any(
        keyword in error_str
        for keyword in ["config", "validation", "missing", "required"]
    ):
        return get_configuration_error(error)
    elif context == "upload":
        return get_upload_error(error, is_onedrive_file)
    else:
        return get_unknown_error(error)
