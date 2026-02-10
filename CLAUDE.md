# CLAUDE.md

Claude Code instructions for this MCP server project.

## Project Overview

SharePoint Document Search MCP (Model Context Protocol) server with Azure AD certificate authentication.
Supports both stdio and HTTP transports for flexible integration.

### Technology Stack

- FastMCP
  - MCP server framework
- Typer
  - CLI interface
- Azure AD certificate authentication
  - Certificate-based authentication for SharePoint
- Quality tools
  - ruff (linting/formatting), ty (type checking)

## Available Skills

Use these skills for common tasks:
- `/quality-check` - Run quality checks and auto-fix code issues
- `/sync-docs` - Update English and Japanese documentation in sync

## Coding Guidelines

**IMPORTANT**: Always run quality checks before committing (use `/quality-check` skill).

### Type Annotations

- Use lowercase built-in types: `list`, `dict`, `set`, `tuple`
- Use pipe syntax for optional types: `| None`

### Import Organization (PEP 8)

- **All imports at file top**: Never use inline imports within functions
- **Import order**: stdlib → third-party → local imports (each group separated by blank line)
- **Avoid duplicate processing**: Function-level imports called repeatedly cause performance overhead
- **Example**:
```python
# Standard library
import logging
from typing import Dict

# Third-party
import typer
from fastmcp import FastMCP

# Local imports
from .server import mcp, setup_logging
```

**❌ Wrong - Function-level imports**:
```python
def some_function():
    from src.module import something  # Never do this
    return something()
```

**✅ Correct - Top-level imports**:
```python
from src.module import something

def some_function():
    return something()
```

## Design Principles

### MCP Server Design Philosophy

**Simplicity over Complexity**
- Prefer simple, explicit APIs over "smart" automatic conversions
- Trust LLM's ability to learn from clear error messages
- Avoid over-engineering: features are easy to add but hard to remove

**YAGNI (You Aren't Gonna Need It)**
- Only implement features when they are clearly needed
- Before adding a feature, ask: "Can LLM handle this on its own?"
- Consider long-term maintenance cost vs. short-term convenience

**Feature Addition Checklist**
- [ ] Is this feature truly necessary? (Can't LLM adapt?)
- [ ] Is there a simpler alternative?
- [ ] Can this be removed in the future without breaking compatibility?
- [ ] Is the maintenance cost acceptable?
- [ ] Does this add significant value to justify the complexity?

**Context Efficiency**
- **LLM Usability First**: Tool descriptions must prioritize LLM's ability to use the tool correctly
- Avoid verbose or redundant phrasing, but never sacrifice clarity for brevity
- Include essential information that LLM needs for correct usage (modes, defaults, behavior)
- Complex tools require detailed descriptions - token cost is acceptable if it ensures correct usage
- Use external docs (README) for human-oriented explanations, not as replacement for tool clarity
- **Lesson learned**: Overly brief descriptions can cause LLM to misuse tools, negating any token savings

### Project-Specific Guidelines

#### MCP Development
- Use proper logging setup to avoid stdout contamination in stdio mode
- Log to stderr only in stdio transport mode
- Follow FastMCP patterns for tool decoration and type hints
- Document tools clearly for LLM consumption

#### SharePoint Integration
- Always validate environment configuration before client initialization
- Handle authentication errors with natural language messages
- Support both file-based and text-based certificate configuration
- Implement proper error handling for SharePoint API calls

#### Response Format Feature
- Support both "detailed" and "compact" response formats for token efficiency
- Always validate response_format parameter with proper fallback to "detailed"
- Use compact format for reduced token usage when full details not needed

#### Error Handling
- Handle transport-specific errors appropriately
- Provide clear error messages for invalid transport selection
- Use proper logging levels for different scenarios
- Implement natural language error messages for better UX

## Project Files Structure

### Core Files
- `src/main.py`
  - CLI entry point with typer
- `src/server.py`
  - MCP server core logic and tool implementations
- `src/config.py`
  - Environment configuration management
- `src/sharepoint_auth.py`
  - Azure AD certificate authentication
- `src/sharepoint_search.py`
  - SharePoint search client implementation
- `src/error_messages.py`
  - Natural language error message handling

### Available MCP Tools
- `sharepoint_docs_search`
  - Search SharePoint documents with keyword queries
  - Support for OneDrive and SharePoint mixed search
  - Multiple search targets (sites, OneDrive folders, or combination)
  - Support for file extension filtering
  - Response format options (detailed/compact)
- `sharepoint_docs_download`
  - Download files from SharePoint using search results

## OneDrive対応機能

### 環境変数
```bash
# OneDriveユーザーとフォルダーの指定
SHAREPOINT_ONEDRIVE_PATHS=user@domain.com[:/folder/path][,user2@domain.com[:/folder/path]]...

# 検索対象の指定（@onedriveキーワード使用）
SHAREPOINT_SITE_NAME=@onedrive,site1,site2
```

### 設定例
```bash
# OneDriveのみ検索
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com,user2@company.com:/Documents/重要書類
SHAREPOINT_SITE_NAME=@onedrive

# OneDriveとSharePointサイトの混合検索
SHAREPOINT_ONEDRIVE_PATHS=manager@company.com:/Documents/経営資料
SHAREPOINT_SITE_NAME=@onedrive,executive-team,board-documents

# 複数SharePointサイトのみ
SHAREPOINT_SITE_NAME=site1,site2,site3
```

### 技術的特徴
- SharePoint REST APIのKQLクエリを使用
- pathフィルターとsiteフィルターの組み合わせ
- メールアドレスからOneDriveパスへの自動変換
- フォルダーレベルまでの詳細指定対応
