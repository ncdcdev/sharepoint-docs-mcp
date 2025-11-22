# Development Guide

This guide covers development-related information including project structure, development commands, and code quality tools.

## Table of Contents

- [Project Structure](#project-structure)
- [Development Commands](#development-commands)
- [Testing Framework](#testing-framework)
- [Code Quality Tools](#code-quality-tools)
- [Debugging](#debugging)

## Project Structure

```
sharepoint-docs-mcp/
├── src/
│   ├── __init__.py
│   ├── server.py            # MCP server core logic
│   ├── main.py              # CLI entry point
│   ├── config.py            # Configuration management
│   ├── sharepoint_auth.py   # Azure AD authentication
│   ├── sharepoint_search.py # SharePoint search client
│   └── error_messages.py    # Error handling
├── tests/
│   ├── __init__.py
│   ├── conftest.py          # Test fixtures and mocks
│   ├── test_config.py       # Configuration tests
│   ├── test_server.py       # Server functionality tests
│   └── test_error_messages.py # Error handling tests
├── docs/                    # Documentation files
├── scripts.py               # Development utility commands
├── pyproject.toml           # Project configuration and dependencies
├── README.md                # English documentation
└── README_ja.md             # Japanese documentation
```

## Development Commands

### Testing

```bash
# Run tests
uv run test

# Run tests with coverage report
uv run test --cov=src --cov-report=html
```

### Code quality checks

```bash
# Lint (static analysis)
uv run lint

# Type checking (ty)
uv run typecheck

# All checks (type checking + lint + tests)
uv run check
```

### Code formatting

```bash
# Format only
uv run fmt

# Auto-fix + format
uv run fix
```

## Testing Framework

- **pytest**: Python testing framework with fixtures and mocking
- **pytest-cov**: Code coverage reporting
- **pytest-mock**: Enhanced mocking capabilities
- 24 unit tests covering core functionality (48% coverage)

## Code Quality Tools

- **ruff**: Fast Python linter and formatter
- **ty**: Fast type checker (pre-release version)

### Configuration Files

- `pyproject.toml`: Project configuration, dependencies, development tool settings
  - pytest configuration: Test discovery and coverage settings
  - ruff configuration: Code style and rule settings
  - ty configuration: Detailed type checking settings

## Debugging

### Using MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### Log Level Adjustment

Detailed logs are output when starting the server. Error details are displayed in standard error output.
