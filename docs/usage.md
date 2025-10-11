# Usage Guide

This guide covers how to use the SharePoint MCP server with various clients and search scenarios.

## Table of Contents

- [MCP Server Startup](#mcp-server-startup)
- [MCP Inspector Verification](#mcp-inspector-verification)
- [Claude Desktop Integration](#claude-desktop-integration)
- [Search Usage Examples](#search-usage-examples)

## MCP Server Startup

### stdio mode (for desktop app integration)
```bash
uv run sharepoint-docs-mcp --transport stdio
```

### HTTP mode (for network services)
```bash
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

### Show help
```bash
uv run sharepoint-docs-mcp --help
```

## MCP Inspector Verification

### stdio mode
1. Open MCP Inspector
2. Select "Command"
3. Command: `uv`
4. Arguments: `run,sharepoint-docs-mcp,--transport,stdio`
5. Working Directory: Project root directory
6. Click "Connect"

### HTTP mode
1. Start server: `uv run sharepoint-docs-mcp --transport http`
2. Select "URL" in MCP Inspector
3. URL: `http://127.0.0.1:8000/mcp/`
4. Click "Connect"

## Claude Desktop Integration

To integrate with Claude Desktop, update the configuration file:

- Windows: `%APPDATA%/Claude/claude_desktop_config.json`
- macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`

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

## Search Usage Examples

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
