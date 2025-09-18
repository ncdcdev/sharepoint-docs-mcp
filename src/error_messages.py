"""
SharePoint検索MCPサーバー用のエラーメッセージ定義
AIエージェント向けに自然言語でわかりやすいエラーメッセージを提供
"""

from enum import Enum


class ErrorCategory(Enum):
    """エラーカテゴリの定義"""

    AUTHENTICATION = "authentication"
    AUTHORIZATION = "authorization"
    NETWORK = "network"
    SEARCH_QUERY = "search_query"
    FILE_NOT_FOUND = "file_not_found"
    CONFIGURATION = "configuration"
    UNKNOWN = "unknown"


class SharePointError(Exception):
    """SharePoint操作用のカスタム例外クラス"""

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
        """AIエージェント向けにフォーマットされたエラーメッセージを取得"""
        return f"{self.message} {self.solution}"


def get_authentication_error(original_error: Exception) -> SharePointError:
    """認証エラーのメッセージを生成"""
    error_str = str(original_error).lower()

    if "certificate" in error_str or "private_key" in error_str:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="証明書または秘密鍵の読み込みに失敗しました。",
            solution="証明書ファイルのパスと秘密鍵ファイルのパス、またはそれらの内容が正しく設定されているか確認してください。",
            original_error=original_error,
        )
    elif "401" in error_str or "unauthorized" in error_str:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="SharePointへの認証に失敗しました。",
            solution="テナントID、クライアントID、証明書の設定を確認し、アプリの登録状況を管理者にお問い合わせください。",
            original_error=original_error,
        )
    else:
        return SharePointError(
            category=ErrorCategory.AUTHENTICATION,
            message="認証処理中にエラーが発生しました。",
            solution="設定を確認するか、管理者にお問い合わせください。",
            original_error=original_error,
        )


def get_authorization_error(original_error: Exception) -> SharePointError:
    """権限エラーのメッセージを生成"""
    return SharePointError(
        category=ErrorCategory.AUTHORIZATION,
        message="SharePointへのアクセスが拒否されました。",
        solution="アプリの権限設定を確認するか、管理者にSharePointサイトへのアクセス権限を依頼してください。",
        original_error=original_error,
    )


def get_network_error(original_error: Exception) -> SharePointError:
    """ネットワークエラーのメッセージを生成"""
    error_str = str(original_error).lower()

    if "timeout" in error_str:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="SharePointサーバーへの接続がタイムアウトしました。",
            solution="ネットワーク接続を確認するか、しばらく時間をおいて再度お試しください。",
            original_error=original_error,
        )
    elif "connection" in error_str:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="SharePointサーバーに接続できませんでした。",
            solution="インターネット接続とサイトURLを確認してください。",
            original_error=original_error,
        )
    else:
        return SharePointError(
            category=ErrorCategory.NETWORK,
            message="ネットワーク通信中にエラーが発生しました。",
            solution="ネットワーク接続を確認してから再度お試しください。",
            original_error=original_error,
        )


def get_search_query_error(original_error: Exception) -> SharePointError:
    """検索クエリエラーのメッセージを生成"""
    return SharePointError(
        category=ErrorCategory.SEARCH_QUERY,
        message="検索クエリの処理中にエラーが発生しました。",
        solution="検索キーワードを変更するか、より具体的な検索条件を指定してください。",
        original_error=original_error,
    )


def get_file_not_found_error(
    file_path: str, original_error: Exception
) -> SharePointError:
    """ファイル不存在エラーのメッセージを生成"""
    return SharePointError(
        category=ErrorCategory.FILE_NOT_FOUND,
        message=f"指定されたファイルが見つかりませんでした: {file_path}",
        solution="ファイルパスが正しいか確認するか、最新の検索結果から正しいパスを取得してください。",
        original_error=original_error,
    )


def get_configuration_error(original_error: Exception) -> SharePointError:
    """設定エラーのメッセージを生成"""
    return SharePointError(
        category=ErrorCategory.CONFIGURATION,
        message="SharePointの設定に問題があります。",
        solution="環境変数の設定を確認し、必要な設定項目がすべて正しく設定されているか確認してください。",
        original_error=original_error,
    )


def get_unknown_error(original_error: Exception) -> SharePointError:
    """不明なエラーのメッセージを生成"""
    return SharePointError(
        category=ErrorCategory.UNKNOWN,
        message="予期しないエラーが発生しました。",
        solution="設定を確認するか、管理者にお問い合わせください。",
        original_error=original_error,
    )


def handle_sharepoint_error(error: Exception, context: str = "") -> SharePointError:
    """
    SharePoint関連のエラーを適切なカテゴリに分類し、自然言語メッセージを生成

    Args:
        error: 発生した例外
        context: エラーが発生したコンテキスト（"auth", "search", "download"など）

    Returns:
        SharePointError: 自然言語化されたエラー
    """
    error_str = str(error).lower()

    # HTTPステータスコードによる分類
    if hasattr(error, "response") and hasattr(error.response, "status_code"):
        status_code = error.response.status_code
        if status_code == 401:
            return get_authentication_error(error)
        elif status_code == 403:
            return get_authorization_error(error)
        elif status_code == 404 and context == "download":
            return get_file_not_found_error("", error)

    # エラーメッセージの内容による分類
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
        return get_file_not_found_error("", error)
    elif any(
        keyword in error_str
        for keyword in ["config", "validation", "missing", "required"]
    ):
        return get_configuration_error(error)
    else:
        return get_unknown_error(error)
