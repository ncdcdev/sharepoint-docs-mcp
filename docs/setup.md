# SharePoint Configuration Setup

This guide covers the detailed setup process for configuring SharePoint authentication with this MCP server.

## Table of Contents

- [Environment Variables Setup](#environment-variables-setup)
- [Certificate Creation](#certificate-creation)
- [Azure AD Application Setup](#azure-ad-application-setup)
- [Tool Description Customization](#tool-description-customization)

## Environment Variables Setup

Create a `.env` file with the following configuration (refer to `.env.example`):

### Common Configuration (Both Authentication Methods)

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

# Search configuration (optional)
SHAREPOINT_DEFAULT_MAX_RESULTS=20
SHAREPOINT_ALLOWED_FILE_EXTENSIONS=pdf,docx,xlsx,pptx,txt,md

# Tool description customization (optional)
# SHAREPOINT_SEARCH_TOOL_DESCRIPTION=Search internal documents
# SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=Download files from search results
```

### Certificate Authentication Configuration (SHAREPOINT_AUTH_MODE=certificate)

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

### OAuth Authentication Configuration (SHAREPOINT_AUTH_MODE=oauth)

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
# Azure AD redirect URI will be: {SERVER_BASE_URL}/auth/callback
# Default: http://localhost:8000
SHAREPOINT_OAUTH_SERVER_BASE_URL=http://localhost:8000

# Allowed MCP client redirect URIs (comma-separated, wildcards supported)
# If not set: All redirect URIs are allowed (convenient for development, not recommended for production)
# If set: Only specified patterns are allowed (recommended for production)
# For local development:
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=http://localhost:*,http://127.0.0.1:*
# For production (e.g., Claude.ai integration):
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=https://claude.ai/*,https://*.anthropic.com/*
```

## Certificate Creation

Create a self-signed certificate for certificate-based authentication:

```bash
mkdir -p cert
openssl genrsa -out cert/private_key.pem 2048
openssl req -new -key cert/private_key.pem -out cert/certificate.csr -subj "/CN=SharePointAuth"
openssl x509 -req -in cert/certificate.csr -signkey cert/private_key.pem -out cert/certificate.pem -days 365
rm cert/certificate.csr
```

Generated files:
- `cert/certificate.pem` - Public certificate (upload to Azure AD)
- `cert/private_key.pem` - Private key (used by server)

## Azure AD Application Setup

Choose the appropriate setup based on your authentication method:

### Option A: Certificate Authentication Setup (Application Permissions)

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

### Option B: OAuth Authentication Setup (User Permissions)

**1. Azure AD Application Registration**
1. Go to [Azure Portal](https://portal.azure.com/) → Entra ID → App registrations
2. Click "New registration"
3. Enter the following:
   - Name: SharePoint MCP OAuth Client
   - Supported account types: Single tenant
   - Redirect URI: Web - `http://localhost:8000/auth/callback`
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

OAuth authentication in this MCP server is handled through **FastMCP's OIDCProxy**, which implements a secure two-layer authentication:

1. **Layer 1 - MCP Client Authentication**:
   - MCP client (e.g., Claude Desktop, MCP Inspector) authenticates with the FastMCP server
   - FastMCP proxy authenticates the user with Microsoft Entra ID
   - Uses PKCE (Proof Key for Code Exchange) for security without client secrets on the client side

2. **Layer 2 - SharePoint API Access**:
   - The authenticated user's token is used to access SharePoint APIs on their behalf
   - User's delegated permissions are used (AllSites.Read/Write)

**Security Features**:
- PKCE (Proof Key for Code Exchange) ensures tokens cannot be intercepted
- Client secrets are only stored on the server side (`.env` file)
- Token validation trusts Azure AD's OAuth flow (no additional JWT verification required for SharePoint tokens)

**Important Notes**:
- Authentication is performed through the MCP client's OAuth flow with Azure AD
- No manual browser login is required - the MCP client handles the OAuth flow automatically
- Tokens are managed and validated by FastMCP (using custom SharePointTokenVerifier)
- The server uses the `/auth/callback` endpoint (FastMCP standard) for OAuth callbacks
- MCP clients can use dynamic ports (e.g., http://localhost:6274/oauth/callback) as FastMCP accepts wildcard localhost URIs

## Tool Description Customization

You can customize MCP tool descriptions in Japanese or other languages:

- `SHAREPOINT_SEARCH_TOOL_DESCRIPTION`: Search tool description (default: "Search for documents in SharePoint")
- `SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION`: Download tool description (default: "Download a file from SharePoint")

Example:
```bash
SHAREPOINT_SEARCH_TOOL_DESCRIPTION=Search internal documents
SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=Download files from search results
```
