import logging
import platform
import sys

from fastmcp import FastMCP

# MCPサーバーインスタンスを作成
mcp = FastMCP(name="SharePointSearchMCP")


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


@mcp.tool
def get_system_info() -> dict[str, str]:
    """
    現在のシステムの基本情報を取得します。

    :return: PythonのバージョンとOSプラットフォームを含む辞書。
    """
    logging.info("Executing get_system_info tool.")
    return {
        "python_version": sys.version,
        "platform": platform.platform(),
    }
