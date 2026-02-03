# Troubleshooting Guide

This guide covers common issues and debugging methods for the SharePoint MCP server.

## Table of Contents

- [Common Issues](#common-issues)
- [Debugging Methods](#debugging-methods)

## Common Issues

### 1. Authentication Errors

```
SharePoint configuration is invalid: SHAREPOINT_TENANT_ID is required
```

**Solutions:**
- Check if `.env` file is configured correctly
- Verify environment variables are loaded properly

### 2. Certificate Errors

```
Certificate file not found: path/to/certificate.pem
```

**Solutions:**
- Verify certificate file path is correct
- Check if certificate is created properly
- Ensure file read permissions are granted

### 3. API Permission Errors

```
Access token request failed
```

**Solutions:**
- Check Azure AD app permission settings
- Verify admin consent has been granted
- Confirm client ID and tenant ID are correct

### 4. Configuration Check Command

```bash
# Check configuration status (using MCP Inspector)
# Execute get_sharepoint_config_status tool
```

### 5. Excel Operations Errors

#### Excel Services Disabled

```
Excel Services is not enabled or not available for this SharePoint site.
```

**Solutions:**
- Request SharePoint administrator to enable Excel Services
- Verify Excel Services is available for the target SharePoint site
- Confirm file is stored in a location where Excel Services is enabled

#### Excel File Not Found

```
The specified Excel file was not found: /sites/team/documents/report.xlsx
```

**Solutions:**
- Verify file path is correct (use `sharepoint_docs_search` to get latest path)
- Check if file has been deleted or moved
- Confirm you have access permissions to the file

#### Sheet Not Found

```
The specified sheet was not found: Sheet2
```

**Solutions:**
- Use `list_sheets` operation to confirm available sheets
- Verify sheet name spelling is correct
- Ensure special characters (like single quotes) are specified correctly

#### Invalid Cell Range

```
The specified cell range is invalid: InvalidRange
```

**Solutions:**
- Verify cell range format is correct (e.g., "Sheet1!A1:C10")
- Confirm sheet name is included in the range specification
- Check if the range is within the actual bounds of the Excel file

## Debugging Methods

### Using MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### Log Level Adjustment

Detailed logs are output when starting the server. Error details are displayed in standard error output.
