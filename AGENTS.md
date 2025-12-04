# xleak - AI Agent Instructions

This file provides AI coding agents with project-specific instructions for working with xleak, an Excel terminal viewer written in Rust.

## Project Overview

xleak is a command-line tool that displays Excel spreadsheets in the terminal with beautiful formatting and export capabilities. It's inspired by [doxx](https://github.com/bgreenwell/doxx) and built with Rust for performance.

**Tech Stack:**
- Language: Rust (Edition 2024)
- Excel Parsing: `calamine` (v0.26)
- CLI Framework: `clap` (v4.5) with derive macros
- Error Handling: `anyhow` (v1.0)
- Terminal Display: `prettytable-rs` (v0.10)
- Interactive TUI: `ratatui` (v0.29) + `crossterm` (v0.28)
- Clipboard: `arboard` (v3.4)
- Date/Time: `chrono` (v0.4)

**Supported File Formats:** `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`

## Project Structure

```
xleak/
├── Cargo.toml              # Dependencies and project configuration
├── src/
│   ├── main.rs            # CLI interface and argument parsing
│   ├── workbook.rs        # Calamine wrapper for reading Excel files
│   ├── tui.rs             # Interactive TUI application state and rendering
│   └── display.rs         # Terminal display and export formatting
├── generate_test_data.py  # Python script to create test Excel files
├── README.md              # User-facing documentation
└── CLAUDE.md              # Context specific to Claude Code
```

## Build and Test Commands

### Building
```bash
# Development build
cargo build

# Release build (optimized)
cargo build --release

# The binary will be at: target/release/xleak
```

### Running
```bash
# Run without building binary
cargo run -- test_data.xlsx

# Run with arguments
cargo run --release -- test_data.xlsx --sheet "Sales" -n 20

# Run compiled binary
./target/release/xleak test_data.xlsx
```

### Installing
```bash
# Install to system (usually ~/.cargo/bin/)
cargo install --path .

# After install, use from anywhere
xleak ~/Documents/report.xlsx
```

### Checking Code Quality
```bash
# Check for compilation errors (no binary produced)
cargo check

# Run clippy for lints
cargo clippy

# Format code
cargo fmt

# Check formatting without modifying
cargo fmt -- --check
```

### Testing
Unit tests are implemented at the bottom of module files (`src/workbook.rs` and `src/tui.rs`).

Run tests with:
```bash
cargo test
```

Test coverage includes:
- Cell value conversions and formatting
- Excel datetime handling
- Column letter conversions
- Cell address parsing

### Creating Test Data
```bash
# Requires Python 3 and openpyxl
pip install openpyxl
python3 generate_test_data.py
```

## Code Style Guidelines

### Rust Conventions
- **Edition:** 2024
- **Formatting:** Use `cargo fmt` with default rustfmt settings
- **Linting:** Address all `cargo clippy` warnings before committing
- **Error Handling:** Use `anyhow::Result<T>` for fallible functions, with `.context()` for error messages
- **Imports:** Group by std → external crates → internal modules

### Module Organization
- `main.rs`: CLI parsing and orchestration only
- `workbook.rs`: All Excel file I/O and data extraction
- `tui.rs`: Interactive TUI state management and event handling
- `display.rs`: All formatting and output (terminal, CSV, JSON, text)

### Naming Conventions
- Functions: `snake_case`
- Types/Structs: `PascalCase`
- Constants: `SCREAMING_SNAKE_CASE`
- Module files: `snake_case.rs`

### Code Comments
- AVOID adding comments unless the code requires explanation
- Focus on "why" not "what"
- Use doc comments (`///`) for public APIs
- Keep inline comments (`//`) minimal

### Pattern Matching
Prefer exhaustive pattern matching over wildcards when handling `CellValue` enum.

## Architecture Guidelines

### Adding New Features

When extending functionality:

1. **New CLI Arguments:** Add to `Cli` struct in `main.rs` with clap derive macros
2. **New Cell Types:** Add to `CellValue` enum in `workbook.rs` and implement `Display` trait
3. **New Export Formats:** Add function to `display.rs` following existing `export_*` pattern
4. **New Data Processing:** Add methods to `SheetData` in `workbook.rs`

### Implemented Features (v0.1.0+)
- ✅ Interactive TUI with ratatui + crossterm
- ✅ Search functionality (full-text with n/N navigation)
- ✅ Formula display mode (view in cell detail popup)
- ✅ Clipboard support (copy cell/row with c/C)
- ✅ Large file optimization with lazy loading (1000+ rows)
- ✅ Multi-sheet navigation (Tab/Shift+Tab)
- ✅ Jump to cell (Ctrl+G for addresses like A100, 10,5)
- ✅ Horizontal scrolling with auto-sized columns (-H flag)

### Future Features (see .planning/ for any relevant documentation)
- 🚧 Cell formatting visualization (colors, borders)
- 🚧 Advanced filtering/sorting
- 🚧 Freeze panes support

### Dependencies

**Current:**
- `calamine`: Excel parsing (do not add alternative parsers)
- `clap`: CLI (use derive macros, not builder pattern)
- `anyhow`: Error handling (prefer over thiserror for applications)
- `prettytable-rs`: Terminal tables for non-interactive mode
- `ratatui` + `crossterm`: Interactive TUI framework
- `arboard`: Cross-platform clipboard support
- `chrono`: Date/time handling for Excel datetime values

**Future (commented out in Cargo.toml):**
- `serde` + `serde_json`: For structured JSON export (currently using manual JSON)

**Do NOT add without discussion:**
- Alternative Excel parsers (stick with calamine)
- Heavy dependencies that increase compile time
- GUI frameworks (this is a terminal tool)

## File Handling

### Excel File Reading
- Use `calamine::open_workbook_auto()` for automatic format detection
- Always check file exists before opening
- Provide helpful error messages with file path context

### Output
- Terminal display: Use prettytable with `FORMAT_BOX_CHARS`
- CSV: Handle quoting for commas and quotes in cell values
- JSON: Manual serialization (no serde yet) with proper escaping
- Text: Tab-separated values

## Error Handling

Use `anyhow` for all error types with contextual messages:

```rust
// Good
wb.load_sheet(&sheet_name)
    .with_context(|| format!("Failed to load sheet '{sheet_name}'"))?;

// Avoid
wb.load_sheet(&sheet_name).unwrap();
```

Provide user-friendly error messages:
- File not found: Show full path
- Sheet not found: List available sheets
- Invalid format: Suggest valid options

## Performance Considerations

- Use `--release` builds for performance testing
- `Cargo.toml` is configured for maximum optimization:
  - `opt-level = 3`
  - `lto = true`
  - `codegen-units = 1`
- Calamine is already fast; avoid loading entire sheets into memory unnecessarily
- For large files, the `-n` limit is critical

## Security

- **No credential handling:** This tool only reads Excel files
- **No network access:** All operations are local
- **Input validation:** Check file exists and format is supported
- **Path traversal:** Use std::path::Path for safe path handling

## Commit Guidelines

### Commit Messages
- Use conventional commits format: `type: description`
- Types: `feat`, `fix`, `docs`, `refactor`, `test`, `chore`
- Keep first line under 72 characters
- Examples:
  - `feat: add TUI mode with ratatui`
  - `fix: handle empty sheets gracefully`
  - `docs: update AGENTS.md with testing section`

### PR Checklist
- [ ] Code compiles: `cargo build --release`
- [ ] No clippy warnings: `cargo clippy`
- [ ] Code is formatted: `cargo fmt`
- [ ] Tested with real Excel files (not just test_data.xlsx)
- [ ] Updated README.md if adding user-facing features
- [ ] Updated AGENTS.md if changing architecture/dependencies

## Development Workflow

1. **Make changes** to source files
2. **Quick check:** `cargo check` (fast, no binary)
3. **Test locally:** `cargo run -- test_data.xlsx`
4. **Run linter:** `cargo clippy`
5. **Format:** `cargo fmt`
6. **Full test:** `cargo build --release` and test with real files
7. **Commit** following commit guidelines above

## Common Tasks for AI Agents

### Adding a new CLI option
1. Add field to `Cli` struct in `src/main.rs`
2. Add clap attribute macros (`#[arg(...)]`)
3. Handle the option in `main()` function
4. Update README.md usage examples

### Adding a new export format
1. Add function `export_<format>()` in `src/display.rs`
2. Add pattern match case in `main.rs` export handling
3. Document in README.md and QUICKSTART.md

### Fixing a display issue
1. Locate issue in `src/display.rs` (likely `display_table()`)
2. Test changes with: `cargo run -- test_data.xlsx`
3. Verify with different data types (numbers, strings, booleans, errors)

### Handling a new cell type
1. Add variant to `CellValue` enum in `src/workbook.rs`
2. Implement `Display` trait for the variant
3. Update `datatype_to_cellvalue()` conversion
4. Update display formatting in `src/display.rs` if needed

## Useful Commands Reference

```bash
# Full development cycle
cargo fmt && cargo clippy && cargo build --release

# Quick iteration
cargo check && cargo run -- test_data.xlsx

# View specific sheet
cargo run -- test_data.xlsx --sheet Employees

# Test exports
cargo run -- test_data.xlsx --export csv > test.csv
cargo run -- test_data.xlsx --export json > test.json

# Performance check (release mode is crucial)
time ./target/release/xleak large_file.xlsx

# Install globally after changes
cargo install --path .
```

## Release Process

xleak uses [cargo-dist](https://github.com/axodotdev/cargo-dist) for automated release builds and distribution. Most installers are automatically published when a version tag is pushed.

See [RELEASE_CHECKLIST.md](./RELEASE_CHECKLIST.md) for the complete step-by-step checklist. You can also create a GitHub issue using the "Release" template to track each release.

### Distribution Channels

**Automated (via cargo-dist + GitHub Actions):**
- ✅ GitHub Releases (binaries, tarballs, installers)
- ✅ Homebrew (bgreenwell/homebrew-tap)
- ✅ Scoop (bgreenwell/scoop-bucket)
- ✅ MSI installer for Windows
- ✅ Shell/PowerShell installers

**Manual:**
- ⚠️ AUR (Arch User Repository) - requires manual update

### Creating a New Release

1. **Update version in Cargo.toml:**
   ```bash
   # Edit version field
   vim Cargo.toml
   ```

2. **Commit and push changes:**
   ```bash
   git add Cargo.toml
   git commit -m "chore: bump version to X.Y.Z"
   git push
   ```

3. **Create and push version tag:**
   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```

4. **GitHub Actions will automatically:**
   - Build binaries for all target platforms
   - Generate installers (shell, PowerShell, MSI, Homebrew, Scoop)
   - Create GitHub Release with release notes
   - Publish Homebrew formula to bgreenwell/homebrew-tap
   - Publish Scoop manifest to bgreenwell/scoop-bucket

5. **Manually publish to AUR:**
   ```bash
   cd ~/Dropbox/devel/xleak

   # Generate PKGBUILD
   cargo aur

   # Get the SHA256 hash from GitHub release
   RELEASE_URL="https://github.com/bgreenwell/xleak/releases/download/vX.Y.Z/xleak-x86_64-unknown-linux-gnu.tar.xz.sha256"
   SHA256=$(curl -sL "$RELEASE_URL" | cut -d' ' -f1)

   # Update PKGBUILD with correct source URL and hash
   cd target/cargo-aur
   # Edit PKGBUILD:
   # - Change source URL to: https://github.com/bgreenwell/xleak/releases/download/vX.Y.Z/xleak-x86_64-unknown-linux-gnu.tar.xz
   # - Update sha256sums with the hash from above

   # Copy to AUR repo
   cp PKGBUILD ~/xleak-bin/

   # Generate .SRCINFO using Docker (macOS doesn't have makepkg)
   docker run --rm -v ~/xleak-bin:/build archlinux:latest /bin/bash -c \
     "useradd -m builder && cd /build && chown -R builder:builder . && \
      su builder -c 'makepkg --printsrcinfo' > .SRCINFO"

   # Commit and push to AUR
   cd ~/xleak-bin
   git add PKGBUILD .SRCINFO
   git commit -m "Update to vX.Y.Z"
   git push origin master
   ```

### Installation Methods (Post-Release)

After a successful release, users can install xleak via:

**macOS/Linux - Homebrew:**
```bash
brew install bgreenwell/tap/xleak
```

**Windows - Scoop:**
```powershell
scoop bucket add bgreenwell https://github.com/bgreenwell/scoop-bucket
scoop install xleak
```

**Windows - MSI Installer:**
Download from: https://github.com/bgreenwell/xleak/releases/latest

**Arch Linux - AUR:**
```bash
yay -S xleak-bin
# or
paru -S xleak-bin
```

**Linux/macOS - Shell Script:**
```bash
curl --proto '=https' --tlsv1.2 -LsSf https://github.com/bgreenwell/xleak/releases/latest/download/xleak-installer.sh | sh
```

### Release Configuration

Release automation is configured in:
- `dist-workspace.toml` - cargo-dist configuration
- `.github/workflows/release.yml` - GitHub Actions workflow
- `Cargo.toml` - Package metadata, WiX GUIDs for MSI
- `wix/main.wxs` - WiX installer template for MSI generation

**Important Notes:**
- Scoop uses the `.zip` archive (not MSI) for portable installation
- AUR uses the `x86_64-unknown-linux-gnu.tar.xz` tarball
- Homebrew formula is auto-generated by cargo-dist
- MSI installer requires WiX toolset (handled by cargo-dist in CI)

### Troubleshooting Releases

**Workflow fails on "out of date files":**
- Run `dist generate` locally and commit the changes
- Check that `allow-dirty = ["ci", "msi"]` is set in dist-workspace.toml

**Scoop installation fails with "file doesn't exist":**
- Verify the manifest points to `.zip` file, not `.msi`
- Check that the SHA256 hash matches the release artifact

**AUR build fails:**
- Verify the source URL points to the correct tarball
- Ensure SHA256 hash is correct
- Test PKGBUILD locally with `makepkg -si` in Docker

## Questions or Issues?

If implementing features:
1. Check .planning/ (local directory, not tracked in git) for any planning documents
2. Review existing code patterns before adding new patterns
3. Keep changes minimal and focused
4. Prefer editing existing files over creating new ones
5. Test with multiple file formats (.xlsx, .xls, .ods)

For TUI implementation (future), reference ratatui tutorial: https://ratatui.rs/tutorial/
