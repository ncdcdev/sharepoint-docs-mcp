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

#### Invalid Excel File Format

```
File is not a valid Excel file or is corrupted
```

**Solutions:**
- Verify the file is a valid .xlsx file (not .xls or other formats)
- Check if the file is corrupted by opening it in Excel locally
- Try re-uploading the file to SharePoint

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
Sheet 'Sheet2' not found. Available sheets: ['Sheet1', 'Summary']
```

**Solutions:**
- First read the file without `sheet` parameter to see all available sheets
- Verify sheet name spelling is correct (case-sensitive)
- Check for leading/trailing spaces in sheet names

#### Invalid Cell Range

```
Invalid cell range: ZZ999999
```

**Solutions:**
- Verify cell range format is correct (e.g., "A1:C10" or "A1")
- Check if the range is within the actual bounds of the Excel file
- Ensure column letters and row numbers are valid

## Debugging Methods

### Using MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### Log Level Adjustment

Detailed logs are output when starting the server. Error details are displayed in standard error output.
