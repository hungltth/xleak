# xleak <img src="assets/logo.jpg" align="right" width="120" />

[![CI](https://img.shields.io/github/actions/workflow/status/bgreenwell/xleak/ci.yml?style=for-the-badge)](https://github.com/bgreenwell/xleak/actions/workflows/ci.yml)
[![Crates.io](https://img.shields.io/crates/v/xleak.svg?style=for-the-badge&color=%23107C41)](https://crates.io/crates/xleak)
[![License: MIT](https://img.shields.io/badge/License-MIT-%232196F3.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)
[![Rust](https://img.shields.io/badge/rust-1.70%2B-%23D34516.svg?style=for-the-badge&logo=rust&logoColor=white)](https://www.rust-lang.org/)

> Expose Excel files in your terminal - no Microsoft Excel required!

Inspired by [doxx](https://github.com/bgreenwell/doxx), `xleak` brings Excel spreadsheets to your command line with beautiful rendering, powerful export capabilities, and a feature-rich interactive TUI.

![xleak demo](assets/demo.gif)

## Features

### Core Functionality
- **Beautiful terminal rendering** with formatted tables
- **Interactive TUI mode** - full keyboard navigation with ratatui
- **Smart data type handling** - numbers right-aligned, text left-aligned, booleans centered
- **Multi-sheet support** - seamlessly navigate between sheets (Tab/Shift+Tab)
- **Multiple export formats** - CSV, JSON, plain text
- **Blazing fast** - powered by `calamine`, the fastest Excel parser in Rust
- **Multiple file formats** - supports `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`

### Interactive TUI Features
- **Full-text search** - search across all cells with `/`, navigate with `n`/`N`
- **Clipboard support** - copy cells (`c`) or entire rows (`C`) to clipboard
- **Formula display** - view Excel formulas in cell detail view (Enter key)
- **Jump to row/column** - press `Ctrl+G` to jump to any cell (e.g., `A100`, `500`, `10,5`)
- **Large file optimization** - lazy loading for files with 1000+ rows
- **Progress indicators** - real-time feedback for long operations
- **Visual cell highlighting** - current row, column, and cell clearly marked

## Installation

### Via Homebrew (macOS/Linux)
```bash
brew tap bgreenwell/xleak
brew install xleak
```

### Via Cargo
```bash
cargo install xleak
```

### Via Nix
```bash
# Run directly
nix run github:bgreenwell/xleak -- file.xlsx

# Install with flakes
nix profile install github:bgreenwell/xleak

# Or enter dev shell
nix develop github:bgreenwell/xleak
```

### Pre-built Binaries
Download pre-built binaries for Windows, Linux, and macOS from the [latest release](https://github.com/bgreenwell/xleak/releases/latest).

### Build from Source
```bash
git clone https://github.com/bgreenwell/xleak.git
cd xleak
cargo install --path .
```

## Usage

### Interactive TUI Mode (Recommended)
```bash
# Launch interactive viewer
xleak quarterly-report.xlsx -i

# Start on a specific sheet
xleak report.xlsx --sheet "Q3 Results" -i

# View formulas by default
xleak data.xlsx -i --formulas
```

**TUI Keyboard Shortcuts:**
- `↑ ↓ ← →` - Navigate cells
- `Enter` - View cell details (including formulas)
- `/` - Search across all cells
- `n` / `N` - Jump to next/previous search result
- `Ctrl+G` - Jump to specific row/cell (e.g., `100`, `A50`, `10,5`)
- `c` - Copy current cell to clipboard
- `C` - Copy entire row to clipboard
- `Tab` / `Shift+Tab` - Switch between sheets
- `?` - Show help
- `q` - Quit

### Non-Interactive Mode

#### View a spreadsheet
```bash
xleak quarterly-report.xlsx
```

#### View a specific sheet
```bash
# By name
xleak report.xlsx --sheet "Q3 Results"

# By index (1-based)
xleak report.xlsx --sheet 2
```

#### Limit displayed rows
```bash
# Show only first 20 rows
xleak large-file.xlsx -n 20

# Show all rows
xleak file.xlsx -n 0
```

#### Export data
```bash
# Export to CSV
xleak data.xlsx --export csv > output.csv

# Export to JSON
xleak data.xlsx --export json > output.json

# Export as plain text (tab-separated)
xleak data.xlsx --export text > output.txt
```

#### Combine options
```bash
# Export specific sheet as CSV
xleak workbook.xlsx --sheet "Sales" --export csv > sales.csv
```

## Examples

```bash
# Launch interactive viewer
xleak quarterly-report.xlsx -i

# Quick preview in non-interactive mode
xleak quarterly-report.xlsx

# See specific sheet with limited rows
xleak financial-data.xlsx --sheet "Summary" -n 10

# Interactive mode with formulas visible
xleak data.xlsx -i --formulas

# Export all data from a sheet
xleak survey-results.xlsx --sheet "Responses" --export csv -n 0
```

## Configuration

xleak supports configuration via a TOML file for persistent settings like default theme and keybindings.

### Config File Location

**Default:** `~/.config/xleak/config.toml` (or `$XDG_CONFIG_HOME/xleak/config.toml`)

**Custom:** Use `--config` flag to specify a different location:
```bash
xleak --config /path/to/config.toml file.xlsx -i
```

### Creating a Config File

1. **Copy the example:**
   ```bash
   cp config.toml.example ~/.config/xleak/config.toml
   ```

2. **Or create manually:**
   ```bash
   mkdir -p ~/.config/xleak
   cat > ~/.config/xleak/config.toml << 'EOF'
   [theme]
   default = "Dracula"

   [ui]
   max_rows = 50
   column_width = 30

   [keybindings]
   profile = "default"
   EOF
   ```

### Configuration Options

#### Theme Settings
```toml
[theme]
# Default theme to use on startup
# Options: "Default", "Dracula", "Solarized Dark", "Solarized Light", "GitHub Dark", "Nord"
default = "Dracula"
```

Press `t` in interactive mode to cycle through themes, or set your favorite as the default.

#### UI Settings
```toml
[ui]
# Default maximum rows in non-interactive mode
max_rows = 50

# Default maximum column width
column_width = 30
```

#### Keybindings
```toml
[keybindings]
# Keybinding profile: "default" or "vim"
profile = "default"

# Custom keybindings (optional)
# [keybindings.custom]
# quit = "q"
# theme_toggle = "t"
# search = "/"
# copy_cell = "c"
```

**VIM Profile:**
Set `profile = "vim"` to use VIM-style navigation:
- `h/j/k/l` - left/down/up/right
- `Ctrl+u/Ctrl+d` - page up/down
- `gg/G` - jump to top/bottom
- `0/$` - jump to row start/end
- `y/Y` - copy cell/row (yank)

See `config.toml.example` for all available options.

## Performance

xleak is optimized for both small and large files:
- **Small files** (< 1000 rows): Instant loading with full eager loading
- **Large files** (≥ 1000 rows): Automatic lazy loading with row caching
  - Memory usage: ~400KB for 10,000 row files
  - Loads only visible rows on demand
  - Progress indicators for long operations

## Comparison to Alternatives

| Tool | Format | Speed | Terminal Native | Interactive | Search | Formulas |
|------|--------|-------|----------------|-------------|--------|----------|
| **xleak** | ✅ xlsx/xls/ods | ⚡ Fast | ✅ Yes | ✅ Full TUI | ✅ Yes | ✅ Yes |
| Excel | ✅ xlsx | ❌ Slow startup | ❌ GUI only | ✅ Yes | ✅ Yes | ✅ Yes |
| pandas | ✅ Many | ❌ Slow | ❌ Python required | ❌ No | ❌ No | ❌ No |
| csvlook | ❌ CSV only | ✅ Fast | ✅ Yes | ❌ No | ❌ No | ❌ No |

## Related Projects

Looking to view Word documents in the terminal? Check out **[doxx](https://github.com/bgreenwell/doxx)** - a terminal viewer for `.docx` files with similar TUI capabilities.

## Built With

- **Rust** - for performance and reliability
- **calamine** - the fastest Excel/ODS parser
- **ratatui** - terminal user interface framework
- **prettytable-rs** - beautiful terminal tables
- **clap** - elegant CLI argument parsing
- **arboard** - cross-platform clipboard support

## Troubleshooting

**"File not found"**
- Ensure the file path is correct
- Use quotes if the filename has spaces: `xleak "My Report.xlsx"`

**"No sheets found"**
- The Excel file might be corrupted
- Try opening it in Excel/LibreOffice first to verify

**"Sheet 'X' not found"**
- Run `xleak file.xlsx` (without --sheet) to see all available sheets
- Sheet names are case-sensitive

## License

MIT

## Credits

- Inspired by [doxx](https://github.com/bgreenwell/doxx) by bgreenwell
- Powered by [calamine](https://github.com/tafia/calamine)
