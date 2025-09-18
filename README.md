# SharePoint Docs MCP Server

> [🇯🇵 日本語版はこちら](README_ja.md)

A Model Context Protocol (MCP) server that provides SharePoint document search functionality.
Supports both stdio and HTTP transports.

Authentication is supported only via Azure AD certificate-based authentication.
Please note that other authentication methods are not supported.

## Features

- SharePoint Search
  - SharePoint document search with certificate authentication
- Certificate Authentication
  - Supports Azure AD certificate-based authentication
- Dual Transport Support
  - Supports both stdio (desktop app integration) and HTTP (network service)
- Proper Logging
  - stderr-based logging configuration to prevent stdout pollution in stdio mode

### SharePoint Features

- sharepoint_docs_search
  - Document search by keywords
  - Response format options (detailed/compact) for token efficiency
- sharepoint_docs_download
  - File download from search results

## Requirements

- Python 3.12 or higher
- uv (package manager)

## Installation

### Option 1: Direct execution with uvx (Recommended)

```bash
# Run directly from GitHub without cloning
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport stdio

# For HTTP mode
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

### Option 2: Development setup

```bash
# Clone the repository
git clone https://github.com/ncdcdev/sharepoint-docs-mcp
cd sharepoint-docs-mcp

# Install dependencies
uv sync --dev
```

## SharePoint Configuration

### 1. Environment Variables Setup

Create a `.env` file with the following configuration (refer to `.env.example`):

```bash
# SharePoint configuration
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_SITE_NAME=yoursite
SHAREPOINT_TENANT_ID=your-tenant-id-here
SHAREPOINT_CLIENT_ID=your-client-id-here

# Leave SHAREPOINT_SITE_NAME empty to search across the entire tenant
# SHAREPOINT_SITE_NAME=

# Certificate authentication configuration (specify either file path or text)
# Priority: 1. Text, 2. File path

# Using file paths
SHAREPOINT_CERTIFICATE_PATH=path/to/your/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/your/private_key.pem

# Or specify directly as text (for Cloud Run etc.)
# Text settings take priority over file paths
# SHAREPOINT_CERTIFICATE_TEXT="-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----"
# SHAREPOINT_PRIVATE_KEY_TEXT="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----"

# Search configuration (optional)
SHAREPOINT_DEFAULT_MAX_RESULTS=20
SHAREPOINT_ALLOWED_FILE_EXTENSIONS=pdf,docx,xlsx,pptx,txt,md

# Tool description customization (optional)
# SHAREPOINT_SEARCH_TOOL_DESCRIPTION=Search internal documents
# SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=Download files from search results
```

### 2. Certificate Creation

Create a self-signed certificate for certificate-based authentication:

```bash
mkdir -p cert
openssl genrsa -out cert/private_key.pem 2048
openssl req -new -key cert/private_key.pem -out cert/certificate.csr -subj "/CN=SharePointAuth"
openssl x509 -req -in cert/certificate.csr -signkey cert/private_key.pem -out cert/certificate.pem -days 365
rm cert/certificate.csr
```

Generated files:
- `cert/certificate.pem`
  - Public certificate (upload to Azure AD)
- `cert/private_key.pem`
  - Private key (used by server)

### 3. Azure AD Certificate Authentication Setup

#### 1. Azure AD Application Registration
1. Go to [Azure Portal](https://portal.azure.com/) → Entra ID → App registrations
2. Click "New registration"
3. Enter application name (e.g., SharePoint MCP Server)
4. Click "Register"

#### 2. Certificate Upload
1. Select the created app → "Certificates & secrets"
2. Click "Upload certificate" in the "Certificates" tab
3. Upload the created `cert/certificate.pem`

#### 3. API Permissions Configuration
1. Go to "API permissions" tab
2. "Add a permission" → "Microsoft Graph" → "Application permissions"
3. Add the following permissions:
   - `Sites.FullControl.All`
     - Full access to SharePoint sites
4. Click "Grant admin consent"

#### 4. Required Information Retrieval
- Tenant ID
  - Directory (tenant) ID from the "Overview" page
- Client ID
  - Application (client) ID from the "Overview" page

### 4. Tool Description Customization (Optional)

You can customize MCP tool descriptions in Japanese or other languages:

- `SHAREPOINT_SEARCH_TOOL_DESCRIPTION`: Search tool description (default: "Search for documents in SharePoint")
- `SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION`: Download tool description (default: "Download a file from SharePoint")

Example:
```bash
SHAREPOINT_SEARCH_TOOL_DESCRIPTION=Search internal documents
SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=Download files from search results
```

## Usage

### MCP Server Startup

**stdio mode (for desktop app integration)**
```bash
uv run sharepoint-docs-mcp --transport stdio
```

**HTTP mode (for network services)**
```bash
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

**Show help**
```bash
uv run sharepoint-docs-mcp --help
```

### MCP Inspector Verification

**stdio mode**
1. Open MCP Inspector
2. Select "Command"
3. Command: `uv`
4. Arguments: `run,sharepoint-docs-mcp,--transport,stdio`
5. Working Directory: Project root directory
6. Click "Connect"

**HTTP mode**
1. Start server: `uv run sharepoint-docs-mcp --transport http`
2. Select "URL" in MCP Inspector
3. URL: `http://127.0.0.1:8000/mcp/`
4. Click "Connect"

### Development Commands

**Testing**
```bash
# Run tests
uv run test

# Run tests with coverage report
uv run test --cov=src --cov-report=html
```

**Code quality checks**
```bash
# Lint (static analysis)
uv run lint

# Type checking (ty)
uv run typecheck

# All checks (type checking + lint + tests)
uv run check
```

**Code formatting**
```bash
# Format only
uv run fmt

# Auto-fix + format
uv run fix
```

## Project Structure

```
sharepoint-docs-mcp/
├── src/
│   ├── __init__.py
│   ├── server.py            # MCP server core logic
│   ├── main.py              # CLI entry point
│   ├── config.py            # Configuration management
│   ├── sharepoint_auth.py   # Azure AD authentication
│   ├── sharepoint_search.py # SharePoint search client
│   └── error_messages.py    # Error handling
├── tests/
│   ├── __init__.py
│   ├── conftest.py          # Test fixtures and mocks
│   ├── test_config.py       # Configuration tests
│   ├── test_server.py       # Server functionality tests
│   └── test_error_messages.py # Error handling tests
├── scripts.py               # Development utility commands
├── pyproject.toml           # Project configuration and dependencies
├── README.md                # English documentation
└── README_ja.md             # Japanese documentation
```

## Claude Desktop Integration

To integrate with Claude Desktop, update the configuration file:

- Windows
  - `%APPDATA%/Claude/claude_desktop_config.json`
- macOS
  - `~/Library/Application\ Support/Claude/claude_desktop_config.json`

### Configuration Example 1: Direct Environment Variables

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uv",
      "args": ["run", "sharepoint-docs-mcp", "--transport", "stdio"],
      "cwd": "/path/to/sharepoint-docs-mcp",
      "env": {
        "SHAREPOINT_BASE_URL": "https://yourcompany.sharepoint.com",
        "SHAREPOINT_SITE_NAME": "yoursite",
        "SHAREPOINT_TENANT_ID": "your-tenant-id-here",
        "SHAREPOINT_CLIENT_ID": "your-client-id-here",
        "SHAREPOINT_CERTIFICATE_PATH": "./cert/certificate.pem",
        "SHAREPOINT_PRIVATE_KEY_PATH": "./cert/private_key.pem"
      }
    }
  }
}
```

### Configuration Example 2: Using .env File (Recommended)

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uv",
      "args": ["run", "sharepoint-docs-mcp", "--transport", "stdio"],
      "cwd": "/path/to/sharepoint-docs-mcp"
    }
  }
}
```

In this case, place the configuration in the `.env` file at the project root.

### Configuration Example 3: Using uvx (No Cloning Required)

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/ncdcdev/sharepoint-docs-mcp", "sharepoint-docs-mcp", "--transport", "stdio"],
      "env": {
        "SHAREPOINT_BASE_URL": "https://yourcompany.sharepoint.com",
        "SHAREPOINT_SITE_NAME": "yoursite",
        "SHAREPOINT_TENANT_ID": "your-tenant-id-here",
        "SHAREPOINT_CLIENT_ID": "your-client-id-here",
        "SHAREPOINT_CERTIFICATE_PATH": "/path/to/certificate.pem",
        "SHAREPOINT_PRIVATE_KEY_PATH": "/path/to/private_key.pem"
      }
    }
  }
}
```

This configuration runs the MCP server directly from GitHub without requiring you to clone the repository locally. Note that `SHAREPOINT_CERTIFICATE_PATH` and `SHAREPOINT_PRIVATE_KEY_PATH` must be absolute paths to your files.

## Development

### Testing Framework

- **pytest**: Python testing framework with fixtures and mocking
- **pytest-cov**: Code coverage reporting
- **pytest-mock**: Enhanced mocking capabilities
- 24 unit tests covering core functionality (48% coverage)

### Code Quality Tools

- **ruff**: Fast Python linter and formatter
- **ty**: Fast type checker (pre-release version)

### Configuration Files

- `pyproject.toml`: Project configuration, dependencies, development tool settings
- pytest configuration: Test discovery and coverage settings
- ruff configuration: Code style and rule settings
- ty configuration: Detailed type checking settings

## Troubleshooting

### Common Issues

#### 1. Authentication Errors
```
SharePoint configuration is invalid: SHAREPOINT_TENANT_ID is required
```
- Check if `.env` file is configured correctly
- Verify environment variables are loaded properly

#### 2. Certificate Errors
```
Certificate file not found: path/to/certificate.pem
```
- Verify certificate file path is correct
- Check if certificate is created properly
- Ensure file read permissions are granted

#### 3. API Permission Errors
```
Access token request failed
```
- Check Azure AD app permission settings
- Verify admin consent has been granted
- Confirm client ID and tenant ID are correct

#### 4. Configuration Check Command
```bash
# Check configuration status (using MCP Inspector)
# Execute get_sharepoint_config_status tool
```

### Debugging Methods

#### Using MCP Inspector
```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

#### Log Level Adjustment
Detailed logs are output when starting the server. Error details are displayed in standard error output.

## License

MIT License - See [LICENSE](LICENSE) file for details.