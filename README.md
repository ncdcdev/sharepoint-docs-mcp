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
  - Support for both SharePoint sites and OneDrive
  - Multiple search targets (sites, OneDrive folders, or mixed)
  - File extension filtering (pdf, docx, xlsx, etc.)
  - Response format options (detailed/compact) for token efficiency
- sharepoint_docs_download
  - File download from search results
  - Automatic method selection for SharePoint vs OneDrive files

### OneDrive Support

This server supports searching both SharePoint sites and OneDrive content with flexible configuration:

- **OneDrive Integration**: Search specific users' OneDrive content
- **Folder-level targeting**: Search specific folders within OneDrive
- **Mixed search**: Combine SharePoint sites and OneDrive in a single search
- **Flexible configuration**: Simple environment variable setup

## Requirements

- Python 3.12
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

You can choose between **Certificate Authentication** (default) and **OAuth Authentication** by setting the `SHAREPOINT_AUTH_MODE` environment variable.

### 1. Environment Variables Setup

Create a `.env` file with the following configuration (refer to `.env.example`):

#### Common Configuration (Both Authentication Methods)

```bash
# SharePoint configuration
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id-here

# Authentication mode ("certificate" or "oauth")
# Default: certificate
SHAREPOINT_AUTH_MODE=certificate

# Search targets (multiple targets supported, comma-separated)
# Options:
#   - @onedrive: Include OneDrive in search (requires SHAREPOINT_ONEDRIVE_PATHS)
#   - @all: Search entire tenant (not recommended for security reasons)
#   - site-name: Specific SharePoint site name(s)
# Examples:
#   - Single site: SHAREPOINT_SITE_NAME=team-site
#   - Multiple sites: SHAREPOINT_SITE_NAME=team-site,project-alpha,hr-docs
#   - OneDrive only: SHAREPOINT_SITE_NAME=@onedrive
#   - Mixed: SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha
SHAREPOINT_SITE_NAME=yoursite

# OneDrive configuration (optional)
# Format: user@domain.com[:/folder/path][,user2@domain.com[:/folder/path]]...
# Examples:
# SHAREPOINT_ONEDRIVE_PATHS=user@company.com,manager@company.com:/Documents/Important
# SHAREPOINT_ONEDRIVE_PATHS=user1@company.com:/Documents/Projects,user2@company.com:/Documents/Archive

# OneDrive configuration (optional)
# Format: user@domain.com[:/folder/path][,user2@domain.com[:/folder/path]]...
# Examples:
# SHAREPOINT_ONEDRIVE_PATHS=user@company.com,manager@company.com:/Documents/Important
# SHAREPOINT_ONEDRIVE_PATHS=user1@company.com:/Documents/Projects,user2@company.com:/Documents/Archive

# Search configuration (optional)
SHAREPOINT_DEFAULT_MAX_RESULTS=20
SHAREPOINT_ALLOWED_FILE_EXTENSIONS=pdf,docx,xlsx,pptx,txt,md

# Tool description customization (optional)
# SHAREPOINT_SEARCH_TOOL_DESCRIPTION=Search internal documents
# SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=Download files from search results
```

#### Certificate Authentication Configuration (SHAREPOINT_AUTH_MODE=certificate)

```bash
# Client ID for certificate authentication
SHAREPOINT_CLIENT_ID=your-client-id-here

# Certificate authentication configuration (specify either file path or text)
# Priority: 1. Text, 2. File path

# Using file paths
SHAREPOINT_CERTIFICATE_PATH=path/to/your/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/your/private_key.pem

# Or specify directly as text (for Cloud Run etc.)
# Text settings take priority over file paths
# SHAREPOINT_CERTIFICATE_TEXT="-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----"
# SHAREPOINT_PRIVATE_KEY_TEXT="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----"
```

#### OAuth Authentication Configuration (SHAREPOINT_AUTH_MODE=oauth)

**Note**: OAuth authentication requires HTTP transport (`--transport http`)

```bash
# OAuth client ID (from Azure AD app registration)
# If not set, falls back to SHAREPOINT_CLIENT_ID
# Typically, only SHAREPOINT_CLIENT_ID is needed for both authentication modes
SHAREPOINT_OAUTH_CLIENT_ID=your-oauth-client-id-here

# OAuth client secret (from Azure AD app registration)
# Required for OAuth mode - create in Azure AD app under "Certificates & secrets"
SHAREPOINT_OAUTH_CLIENT_SECRET=your-oauth-client-secret-here

# FastMCP server base URL (for OAuth callbacks)
# Default: http://localhost:8000
SHAREPOINT_OAUTH_SERVER_BASE_URL=http://localhost:8000

# OAuth redirect URI (must match Azure AD app configuration)
# Default: http://localhost:8000/auth/callback
# Note: Changed from /oauth/callback to /auth/callback (FastMCP standard)
SHAREPOINT_OAUTH_REDIRECT_URI=http://localhost:8000/auth/callback
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

### 3. Azure AD Application Setup

Choose the appropriate setup based on your authentication method:

#### Option A: Certificate Authentication Setup (Application Permissions)

**1. Azure AD Application Registration**
1. Go to [Azure Portal](https://portal.azure.com/) → Entra ID → App registrations
2. Click "New registration"
3. Enter application name (e.g., SharePoint MCP Server)
4. Click "Register"

**2. Certificate Upload**
1. Select the created app → "Certificates & secrets"
2. Click "Upload certificate" in the "Certificates" tab
3. Upload the created `cert/certificate.pem`

**3. API Permissions Configuration**
1. Go to "API permissions" tab
2. "Add a permission" → "Microsoft Graph" → "Application permissions"
3. Add the following permissions:
   - `Sites.FullControl.All` - Full access to SharePoint sites
4. Click "Grant admin consent"

**4. Required Information**
- Tenant ID: Directory (tenant) ID from the "Overview" page
- Client ID: Application (client) ID from the "Overview" page

#### Option B: OAuth Authentication Setup (User Permissions)

**1. Azure AD Application Registration**
1. Go to [Azure Portal](https://portal.azure.com/) → Entra ID → App registrations
2. Click "New registration"
3. Enter the following:
   - Name: SharePoint MCP OAuth Client
   - Supported account types: Single tenant
   - Redirect URI: Web - `http://localhost:8000/auth/callback` (Note: Changed from /oauth/callback)
4. Click "Register"

**2. Client Secret Configuration**
1. Select the created app → "Certificates & secrets"
2. Click "New client secret"
3. Add description (e.g., "MCP Server Secret")
4. Set expiration period (e.g., 24 months)
5. Click "Add"
6. **Important**: Copy the secret value immediately (it won't be shown again)
7. Save this value to `SHAREPOINT_OAUTH_CLIENT_SECRET` environment variable

**3. Authentication Configuration**
1. Select the created app → "Authentication"
2. Under "Platform configurations", verify the redirect URI is set to `http://localhost:8000/auth/callback`
3. Under "Advanced settings":
   - Allow public client flows: No
4. Save changes

**4. API Permissions Configuration (Delegated Permissions)**
1. Go to "API permissions" tab
2. "Add a permission" → "SharePoint" → "Delegated permissions"
3. Add the following permissions:
   - `AllSites.Read` - Read items in all site collections
   - `AllSites.Write` - Read and write items in all site collections (if file downloads are needed)
   - `User.Read` - Read user profile (automatically added)
4. Click "Grant admin consent" (admin consent required)

**5. Required Information**
- Tenant ID: Directory (tenant) ID from the "Overview" page
- OAuth Client ID: Application (client) ID from the "Overview" page
- OAuth Client Secret: Secret value from step 2

**6. Authentication Flow**

OAuth authentication in this MCP server is handled through **FastMCP's AzureProvider**, which implements a secure two-layer authentication:

1. **MCP Client Authentication**: When an MCP client (e.g., Claude Desktop, MCP Inspector) connects, it authenticates with Microsoft Entra ID
2. **SharePoint Access**: The authenticated user's token is used to access SharePoint APIs on their behalf

**Important Notes**:
- Authentication is performed through the MCP client's OAuth flow
- No manual browser login is required - the MCP client handles the OAuth flow automatically
- Tokens are managed by FastMCP and cached securely
- The server uses the `/auth/callback` endpoint (FastMCP standard) for OAuth callbacks

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

## Usage Examples

### SharePoint Site Search Only
```bash
# Search specific SharePoint site
SHAREPOINT_SITE_NAME=team-site

# Search multiple SharePoint sites
SHAREPOINT_SITE_NAME=team-site,project-alpha,hr-docs
```

### OneDrive Search Only
```bash
# Search specific users' OneDrive (entire OneDrive)
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com,user2@company.com
SHAREPOINT_SITE_NAME=@onedrive

# Search specific folders in OneDrive
SHAREPOINT_ONEDRIVE_PATHS=manager@company.com:/Documents/Important,user@company.com:/Documents/Projects
SHAREPOINT_SITE_NAME=@onedrive
```

### Mixed Search (OneDrive + SharePoint)
```bash
# Search OneDrive and SharePoint sites together
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com:/Documents/Projects,manager@company.com:/Documents/Important
SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha
```

### Common Use Cases

**Executive Team Setup**
```bash
# Search executive OneDrive folders and board documents
SHAREPOINT_ONEDRIVE_PATHS=ceo@company.com:/Documents/Executive,cfo@company.com:/Documents/Finance
SHAREPOINT_SITE_NAME=@onedrive,executive-team,board-documents
```

**Project Team Setup**
```bash
# Search project members' work folders and team sites
SHAREPOINT_ONEDRIVE_PATHS=pm@company.com:/Documents/ProjectA,dev@company.com:/Documents/ProjectA
SHAREPOINT_SITE_NAME=@onedrive,project-a-team,project-a-docs
```

**Sales Team Setup**
```bash
# Search sales OneDrive folders and customer sites
SHAREPOINT_ONEDRIVE_PATHS=sales1@company.com:/Documents/Customers,sales2@company.com:/Documents/Proposals
SHAREPOINT_SITE_NAME=@onedrive,sales-team,customer-portal
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