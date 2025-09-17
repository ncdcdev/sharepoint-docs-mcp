import logging
import platform
import sys
from typing import Any

from fastmcp import FastMCP

from .config import config
from .sharepoint_auth import SharePointCertificateAuth
from .sharepoint_search import SharePointSearchClient

# MCPサーバーインスタンスを作成
mcp = FastMCP(name="SharePointSearchMCP")

# SharePointクライアントのグローバルインスタンス
_sharepoint_client: SharePointSearchClient | None = None


def setup_logging():
    """
    すべてのログ出力をstderrに向けるロギングを設定します。
    これにより、stdioトランスポートのstdoutが汚染されるのを防ぎます。
    """
    log_formatter = logging.Formatter("%(asctime)s [%(levelname)s] - %(message)s")
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # stdoutへの出力を防ぐため、既存のハンドラをクリア
    root_logger.handlers.clear()

    # stderrにログを出力するハンドラを追加
    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setFormatter(log_formatter)
    root_logger.addHandler(stream_handler)

    logging.info("Logging configured to output to stderr.")


def _get_sharepoint_client() -> SharePointSearchClient:
    """SharePointクライアントを取得または初期化"""
    global _sharepoint_client

    if _sharepoint_client is None:
        # 設定の検証
        validation_errors = config.validate()
        if validation_errors:
            error_msg = "SharePoint configuration is invalid: " + "; ".join(
                validation_errors
            )
            logging.error(error_msg)
            raise ValueError(error_msg)

        # 認証クライアントを初期化
        auth = SharePointCertificateAuth(
            tenant_id=config.tenant_id,
            client_id=config.client_id,
            certificate_path=config.certificate_path,
            private_key_path=config.private_key_path,
        )

        # SharePointクライアントを初期化
        _sharepoint_client = SharePointSearchClient(
            site_url=config.site_url,
            auth=auth,
        )

        logging.info("SharePoint client initialized successfully")

    return _sharepoint_client


@mcp.tool
def search_sharepoint_documents(
    query: str,
    max_results: int = 20,
    file_extensions: list[str] | None = None,
) -> list[dict[str, Any]]:
    """
    SharePointでドキュメントを検索します。

    Args:
        query: 検索クエリ（キーワード）
        max_results: 返す最大結果数（デフォルト: 20）
        file_extensions: 検索対象のファイル拡張子のリスト（例: ["pdf", "docx"]）

    Returns:
        検索結果のリスト。各結果は以下のキーを含む辞書：
        - title: ドキュメントのタイトル
        - path: ドキュメントのパス
        - size: ファイルサイズ（バイト）
        - last_modified: 最終更新日時
        - file_extension: ファイル拡張子
        - summary: ハイライトされた要約
        - author: 作成者
    """
    logging.info(f"Searching SharePoint documents with query: '{query}'")

    try:
        client = _get_sharepoint_client()

        # ファイル拡張子のフィルタリング
        if file_extensions:
            # 設定で許可された拡張子のみを使用
            allowed_extensions = [
                ext
                for ext in file_extensions
                if ext.lower() in config.allowed_file_extensions
            ]
            if not allowed_extensions:
                logging.warning("No allowed file extensions found in the request")
        else:
            allowed_extensions = None

        # 最大結果数の制限
        max_results = min(max_results, 100)  # 最大100件に制限

        # 検索実行
        results = client.search_documents(
            query=query,
            max_results=max_results,
            file_extensions=allowed_extensions,
        )

        logging.info(f"SharePoint search completed. Found {len(results)} documents")
        return results

    except Exception as e:
        error_msg = f"SharePoint search failed: {str(e)}"
        logging.error(error_msg)
        raise RuntimeError(error_msg) from e


@mcp.tool
def get_system_info() -> dict[str, str]:
    """
    現在のシステムの基本情報を取得します。

    Returns:
        PythonのバージョンとOSプラットフォームを含む辞書。
    """
    logging.info("Executing get_system_info tool.")
    return {
        "python_version": sys.version,
        "platform": platform.platform(),
    }
