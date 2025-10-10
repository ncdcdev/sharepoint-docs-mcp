import logging

import typer

from .config import config
from .server import mcp, register_tools, setup_logging

# typerアプリケーションを作成
app = typer.Typer()


@app.command()
def main(
    transport: str = typer.Option(
        "stdio",
        "--transport",
        help="使用するトランスポートプロトコル ('stdio' または 'http')。",
    ),
    host: str = typer.Option(
        "127.0.0.1", "--host", help="HTTPサーバーのホスト（httpモードのみ）。"
    ),
    port: int = typer.Option(
        8000, "--port", help="HTTPサーバーのポート（httpモードのみ）。"
    ),
):
    """
    stdioまたはhttpトランスポートでMCPサーバーを起動します。
    """
    setup_logging()
    register_tools()

    # OAuth認証モードのバリデーション
    if config.is_oauth_mode and transport == "stdio":
        logging.error(
            "OAuth authentication mode is only supported with 'http' transport. "
            "Please use --transport http or switch to certificate authentication mode."
        )
        raise typer.Exit(code=1)

    if transport == "stdio":
        logging.info("Starting MCP server with stdio transport...")
        mcp.run()  # transport='stdio'がデフォルト
    elif transport == "http":
        logging.info(f"Starting MCP server with http transport on {host}:{port}...")

        # OAuthモードの情報を表示
        if config.is_oauth_mode:
            logging.info(
                "OAuth authentication mode enabled. "
                "FastMCP AzureProvider handles authentication via /auth/callback endpoint."
            )

        # HTTPサーバー起動
        mcp.run(transport="http", host=host, port=port)
    else:
        logging.error(f"Invalid transport: {transport}. Please use 'stdio' or 'http'.")
        raise typer.Exit(code=1)


if __name__ == "__main__":
    app()
