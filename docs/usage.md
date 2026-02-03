# Usage Guide

This guide covers how to use the SharePoint MCP server with various clients and search scenarios.

## Table of Contents

- [MCP Server Startup](#mcp-server-startup)
- [MCP Inspector Verification](#mcp-inspector-verification)
- [Claude Desktop Integration](#claude-desktop-integration)
- [Search Usage Examples](#search-usage-examples)
- [Excel Operations Usage Examples](#excel-operations-usage-examples)

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

## Excel Operations Usage Examples

You can perform operations on Excel files in SharePoint: list sheets, get sheet images, and retrieve cell range data.

### Prerequisites

- SharePoint Excel Services must be enabled
- Excel files must be stored in a SharePoint library
- Appropriate access permissions required

### Basic Workflow

1. **Search for Excel Files**
```python
# Use sharepoint_docs_search tool
results = sharepoint_docs_search(
    query="budget",
    file_extensions=["xlsx"]
)
# Get file_path from results
file_path = results[0]["path"]
# Example: "/sites/finance/Shared Documents/budget_2024.xlsx"
```

2. **List Sheets**
```python
# Use sharepoint_excel_operations tool
sheets_xml = sharepoint_excel_operations(
    operation="list_sheets",
    file_path=file_path
)
# Returns XML format sheet list
# Identify the sheet name you need
```

3. **Get Sheet Image**
```python
# Get visual preview of a specific sheet
image_base64 = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="Sheet1"
)
# Returns base64-encoded image data
# Can be saved or displayed as an image
```

4. **Get Cell Range Data**
```python
# Get data from a specific cell range
range_xml = sharepoint_excel_operations(
    operation="get_range",
    file_path=file_path,
    range_spec="Sheet1!A1:D10"
)
# Returns XML format cell data
# Can be used for data analysis or report generation
```

### Operation Types

#### list_sheets
List all sheets in XML format.

**Parameters:**
- `operation`: "list_sheets"
- `file_path`: Path to Excel file (from search results)

**Returns:** XML format sheet list

**Example:**
```python
sheets = sharepoint_excel_operations(
    operation="list_sheets",
    file_path="/sites/team/Shared Documents/report.xlsx"
)
```

#### get_image
Get a screenshot of the specified sheet in base64 format.

**Parameters:**
- `operation`: "get_image"
- `file_path`: Path to Excel file
- `sheet_name`: Sheet name (required)

**Returns:** base64-encoded image data (PNG format)

**Example:**
```python
image = sharepoint_excel_operations(
    operation="get_image",
    file_path="/sites/team/Shared Documents/report.xlsx",
    sheet_name="Summary"
)
# Save as image
import base64
with open("sheet_preview.png", "wb") as f:
    f.write(base64.b64decode(image))
```

#### get_range
Get data from the specified cell range in XML format.

**Parameters:**
- `operation`: "get_range"
- `file_path`: Path to Excel file
- `range_spec`: Cell range (required, e.g., "Sheet1!A1:C10")

**Returns:** XML format cell data

**Example:**
```python
data = sharepoint_excel_operations(
    operation="get_range",
    file_path="/sites/team/Shared Documents/report.xlsx",
    range_spec="Sheet1!A1:E20"
)
```

### Handling Special Characters

Single quotes (') in sheet names or cell ranges are automatically escaped.

**Example:**
```python
# Specify sheet name "John's Report"
image = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="John's Report"  # Automatically escaped to "John''s Report"
)
```

### Common Use Cases

**Budget Data Analysis**
```python
# 1. Search for budget file
results = sharepoint_docs_search(query="budget 2024", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. Check sheet list
sheets = sharepoint_excel_operations(operation="list_sheets", file_path=file_path)

# 3. Get budget data
budget_data = sharepoint_excel_operations(
    operation="get_range",
    file_path=file_path,
    range_spec="Budget!A1:F100"
)
```

**Report Visual Preview**
```python
# 1. Search for report file
results = sharepoint_docs_search(query="monthly report", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. Get summary sheet image
summary_image = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="Summary"
)

# 3. Save or display image
import base64
with open("monthly_summary.png", "wb") as f:
    f.write(base64.b64decode(summary_image))
```
