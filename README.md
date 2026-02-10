# SharePoint Docs MCP Server

> [🇯🇵 日本語版はこちら](README_ja.md)

A Model Context Protocol (MCP) server that provides SharePoint document search functionality.
Supports both stdio and HTTP transports.

## Authentication Methods

Two authentication methods are supported:

- **Certificate Authentication** (Application Permissions)
  - Uses Azure AD certificate-based authentication
  - Supports both stdio and HTTP transports
  - Recommended for server applications and automation
- **OAuth Authentication** (User Permissions)
  - Uses OAuth 2.0 Authorization Code Flow with PKCE
  - HTTP transport only (browser-based authentication required)
  - Recommended for user-delegated access scenarios

## Features

### SharePoint Features

- **sharepoint_docs_search**
  - Document search by keywords
  - Support for both SharePoint sites and OneDrive
  - Multiple search targets (sites, OneDrive folders, or mixed)
  - File extension filtering (pdf, docx, xlsx, etc.)
  - Response format options (detailed/compact) for token efficiency
- **sharepoint_docs_download**
  - File download from search results
  - Automatic method selection for SharePoint vs OneDrive files
- **sharepoint_excel**
  - Read or search Excel files in SharePoint
  - Search mode: find cells containing specific text with `query` parameter
  - Read mode: get data from specific sheets/ranges with `sheet` and `cell_range` parameters
  - **Automatic header inclusion**: when `cell_range` is specified, frozen rows (headers) are automatically included by default
    - Set `include_frozen_rows=False` to get only the specified range
    - For sheets with `frozen_rows=0`, use `expand_axis_range=True` to include row 1 (for columns) or column A (for rows)
  - **Cell style information** (optional): set `include_cell_styles=True` to get background colors, column widths, and row heights
    - Default is `False` to minimize token usage
    - Useful for identifying highlighted cells, colored headers, or visually emphasized content
  - Response includes cell data in `rows` (value and coordinate) and structural information when available
  - Structural info: sheet name, dimensions, frozen_rows, frozen_cols, freeze_panes (when present), merged_ranges (when merged cells exist)
  - No Excel Services dependency - uses direct file download + openpyxl parsing

### OneDrive Support

This server supports searching both SharePoint sites and OneDrive content with flexible configuration:

- **OneDrive Integration**: Search specific users' OneDrive content
- **Folder-level targeting**: Search specific folders within OneDrive
- **Mixed search**: Combine SharePoint sites and OneDrive in a single search
- **Flexible configuration**: Simple environment variable setup

## Requirements

- Python 3.12
- uv (package manager)

## Quick Start

### 1. Installation

```bash
# Run directly from GitHub without cloning
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport stdio
```

### 2. Configuration

Create a `.env` file with your SharePoint credentials:

```bash
# Basic configuration
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id-here
SHAREPOINT_CLIENT_ID=your-client-id-here
SHAREPOINT_SITE_NAME=yoursite

# For certificate authentication
SHAREPOINT_CERTIFICATE_PATH=path/to/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/private_key.pem

# For OAuth authentication (HTTP transport only)
# SHAREPOINT_AUTH_MODE=oauth
# SHAREPOINT_OAUTH_CLIENT_SECRET=your-oauth-client-secret-here
# SHAREPOINT_OAUTH_SERVER_BASE_URL=https://your-server.com
# If not set: all URIs allowed (dev). If set: only specified patterns (production recommended)
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=https://claude.ai/*,https://*.anthropic.com/*
```

See [Setup Guide](docs/setup.md) for detailed configuration instructions.

### 3. Run the Server

```bash
# stdio mode (for Claude Desktop)
uv run sharepoint-docs-mcp --transport stdio

# HTTP mode (for network services)
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

## Documentation

- 📘 [Setup Guide](docs/setup.md) - Detailed Azure AD and environment configuration
- 📗 [Usage Guide](docs/usage.md) - MCP client integration and search examples
- 📙 [Development Guide](docs/development.md) - Project structure and development commands
- 📕 [Troubleshooting Guide](docs/troubleshooting.md) - Common issues and debugging

## License

MIT License - See [LICENSE](LICENSE) file for details.
