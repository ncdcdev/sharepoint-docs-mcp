# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.5.0] - 2026-02-11

### Added
- Cell style information retrieval with `include_cell_styles` parameter (#51)
- Automatic frozen row header detection with `include_frozen_rows` parameter (#48)
- Optional axis range expansion with `expand_axis_range` parameter (default: false) (#47, #53)
- Metadata-only mode with `metadata_only` parameter for efficient Excel file inspection

### Changed
- Improved token efficiency by omitting redundant Excel response fields (#44, #52)
- Enhanced header detection logic to prevent excessive data reads
- Optimized cell style retrieval with caching mechanism (#51)

### Fixed
- Refactored header detection to reduce code duplication (#58)
- Fixed freeze_panes scroll position bug
- Improved merged cell handling and cache efficiency (#35, #39, #40)
- Strengthened security validation for Excel operations

### Performance
- Eliminated duplicate range normalization in merged cell cache (#39, #40)
- Optimized test performance and reduced warning logs (#52)

### Documentation
- Updated MCP tool description guidelines based on lessons learned
- Clarified parameter behavior and validation rules
- Improved docstrings for better API understanding

## [0.4.0] - Previous Release

Initial release with SharePoint Excel operations support.

[0.5.0]: https://github.com/ncdcdev/sharepoint-docs-mcp/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/ncdcdev/sharepoint-docs-mcp/releases/tag/v0.4.0
