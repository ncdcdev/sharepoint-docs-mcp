# Troubleshooting Guide

This guide covers common issues for the SharePoint MCP server.

## Table of Contents

- [Common Issues](#common-issues)

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
