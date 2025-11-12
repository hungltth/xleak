# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- Double keypress issue on Windows by filtering key release events (thanks [@clindholm](https://github.com/clindholm)! [#2](https://github.com/bgreenwell/xleak/issues/2))

## [0.1.0] - 2025-01-08

### Added
- Initial release of xleak
- Interactive TUI mode with ratatui
- Support for multiple Excel formats (.xlsx, .xls, .xlsm, .xlsb, .ods)
- Search functionality across sheets
- Formula display mode
- Export to CSV, JSON, and text formats
- Lazy loading for large files
- Sheet selection
- Row limit option
- Cross-platform support (Linux, macOS, Windows)

[Unreleased]: https://github.com/greenwbm/xleak/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/greenwbm/xleak/releases/tag/v0.1.0
