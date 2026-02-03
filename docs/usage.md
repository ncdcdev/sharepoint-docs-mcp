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

You can parse Excel files in SharePoint and retrieve data in JSON format. By default, returns lightweight response with value and coordinate only. Use `include_formatting=true` for additional formatting information.

### Prerequisites

- Excel files must be stored in a SharePoint library or OneDrive
- Appropriate access permissions required
- No Excel Services dependency

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

2. **Parse Excel File to JSON (Default: Lightweight)**
```python
# Use sharepoint_excel_to_json tool (default: lightweight response)
json_data = sharepoint_excel_to_json(file_path=file_path)

# Parse JSON response
import json
data = json.loads(json_data)

# Access sheet information
for sheet in data["sheets"]:
    print(f"Sheet: {sheet['name']}")
    print(f"Dimensions: {sheet['dimensions']}")

    # Access cell data (value and coordinate only)
    for row in sheet["rows"]:
        for cell in row:
            print(f"{cell['coordinate']}: {cell['value']}")
```

3. **Parse Excel File with Formatting (Optional)**
```python
# Use include_formatting=true for additional information
json_data = sharepoint_excel_to_json(
    file_path=file_path,
    include_formatting=True
)

# Parse JSON response
data = json.loads(json_data)

# Access cell data with formatting
for sheet in data["sheets"]:
    for row in sheet["rows"]:
        for cell in row:
            print(f"{cell['coordinate']}: {cell['value']}")
            if "fill" in cell:
                print(f"  Fill color: {cell['fill']['fg_color']}")
            if "merged" in cell:
                print(f"  Merged range: {cell['merged']['range']}")
```

### JSON Output Format

#### Default Format (Lightweight)

By default, returns only essential cell information for optimal performance:

```json
{
  "file_path": "/sites/test/Shared Documents/budget.xlsx",
  "sheets": [
    {
      "name": "Summary",
      "dimensions": "A1:E10",
      "rows": [
        [
          {
            "value": "Department",
            "coordinate": "A1"
          },
          {
            "value": 12500,
            "coordinate": "B1"
          }
        ]
      ]
    }
  ]
}
```

#### With Formatting (include_formatting=true)

When `include_formatting=true`, includes additional formatting information:

```json
{
  "file_path": "/sites/test/Shared Documents/budget.xlsx",
  "sheets": [
    {
      "name": "Summary",
      "dimensions": "A1:E10",
      "rows": [
        [
          {
            "value": "Department",
            "coordinate": "A1",
            "data_type": "s",
            "fill": {
              "pattern_type": "solid",
              "fg_color": "#CCCCCC",
              "bg_color": null
            },
            "merged": {
              "range": "A1:B1",
              "is_top_left": true
            },
            "width": 15.0,
            "height": 20.0
          }
        ]
      ]
    }
  ]
}
```

### Available Cell Information

**Default (always included):**
- **value**: Cell value (string, number, date, formula, etc.)
- **coordinate**: Cell position (e.g., "A1", "B2")

**With include_formatting=true:**
- **data_type**: Data type code (`s`=string, `n`=number, `f`=formula, etc.)
- **fill**: Fill color information (pattern type, foreground/background colors)
- **merged**: Merged cell information (range, position)
- **width**: Column width
- **height**: Row height

### Common Use Cases

**Extract All Budget Data**
```python
# 1. Search for budget file
results = sharepoint_docs_search(query="budget 2024", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. Get all Excel data as JSON
json_data = sharepoint_excel_to_json(file_path=file_path)
data = json.loads(json_data)

# 3. Process specific sheet
for sheet in data["sheets"]:
    if sheet["name"] == "Budget":
        for row in sheet["rows"]:
            # Extract values from each cell
            values = [cell["value"] for cell in row]
            print(values)
```

**Analyze Cell Formatting**
```python
# 1. Get Excel data with formatting
json_data = sharepoint_excel_to_json(file_path=file_path, include_formatting=True)
data = json.loads(json_data)

# 2. Find cells with specific formatting
for sheet in data["sheets"]:
    for row in sheet["rows"]:
        for cell in row:
            # Find colored cells
            if cell.get("fill", {}).get("fg_color"):
                print(f"Colored cell at {cell['coordinate']}: {cell['value']}")
                print(f"  Color: {cell['fill']['fg_color']}")
```

**Export to Different Format**
```python
# 1. Get Excel data
json_data = sharepoint_excel_to_json(file_path=file_path)
data = json.loads(json_data)

# 2. Convert to CSV-like format
import csv
for sheet in data["sheets"]:
    with open(f"{sheet['name']}.csv", "w", newline="") as f:
        writer = csv.writer(f)
        for row in sheet["rows"]:
            values = [cell["value"] if cell["value"] is not None else "" for cell in row]
            writer.writerow(values)
```

**Process Multiple Sheets**
```python
# 1. Get all Excel data
json_data = sharepoint_excel_to_json(file_path=file_path)
data = json.loads(json_data)

# 2. Process each sheet
summary = {}
for sheet in data["sheets"]:
    sheet_name = sheet["name"]
    row_count = len(sheet["rows"])
    col_count = len(sheet["rows"][0]) if sheet["rows"] else 0

    summary[sheet_name] = {
        "dimensions": sheet["dimensions"],
        "rows": row_count,
        "columns": col_count
    }

print(json.dumps(summary, indent=2))
```
