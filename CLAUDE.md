# CLAUDE.md

Claude Code instructions for this MCP server project.

## Project Overview

A SharePoint Search MCP (Model Context Protocol) server that supports both stdio and HTTP transports.

### Technology Stack

- **FastMCP**: MCP server framework
- **Typer**: CLI interface
- **Tools**: ruff, ty

## Development Commands

### Setup
```bash
uv sync --dev           # Install dependencies
uv run sharepoint-search-mcp --transport stdio   # Start stdio mode
uv run sharepoint-search-mcp --transport http    # Start HTTP mode
```

### Quality Checks
```bash
uv run check           # Run all checks (typecheck + lint) - verification only
uv run fix             # Auto-fix + format code (includes --unsafe-fixes)
uv run typecheck       # Type checking with ty
uv run lint            # Lint code with ruff
uv run fmt             # Format code with ruff
```

## Coding Guidelines

**IMPORTANT**: Always run quality checks before committing:
```bash
uv run check     # Required - runs all quality checks (verification only)
uv run fix       # Auto-fix and format code (includes --unsafe-fixes)
```

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

### MCP Development

- Use proper logging setup to avoid stdout contamination in stdio mode
- Log to stderr only in stdio transport mode
- Follow FastMCP patterns for tool decoration and type hints
- Document tools clearly for LLM consumption

### Error Handling

- Handle transport-specific errors appropriately
- Provide clear error messages for invalid transport selection
- Use proper logging levels for different scenarios

### 日本語文章でのMarkdownフォーマット

日本語でドキュメントを作成する際は、以下のガイドラインに従う：

#### 太字の使用を避ける
**❌ Wrong**:
```markdown
- **機能名**: 機能の説明
- **設定項目**: 設定の説明
```

**✅ Correct**:
```markdown
- 機能名
  - 機能の説明
- 設定項目
  - 設定の説明
```

#### コロン「:」の使用を最小限にする
文末のコロンは日本語として不自然なため、本当に必要な場合のみ使用する

**❌ Wrong**:
```markdown
以下の設定を行います：
必要な情報の取得：
```

**✅ Correct**:
```markdown
以下の設定を行います
必要な情報の取得
```

#### 箇条書きでの構造化
見出しと説明を区別する際は、説明部分を1段階深くインデントする

**❌ Wrong**:
```markdown
- 項目名: 項目の説明
- 別の項目名: 別の項目の説明
```

**✅ Correct**:
```markdown
- 項目名
  - 項目の説明
- 別の項目名
  - 別の項目の説明
```