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