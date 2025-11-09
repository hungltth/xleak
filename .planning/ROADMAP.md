# xleak Roadmap

## v0.1.0 Accomplishments ✅

The initial release includes:
- Interactive TUI with ratatui + crossterm
- Full keyboard navigation (arrows, Page Up/Down, Home/End, vim-style jumps)
- Search functionality (/, n, N for next/previous match)
- Formula viewing (Enter to see cell details including formulas)
- Clipboard support (c for cell, C for row)
- Multi-sheet navigation (Tab/Shift+Tab)
- Jump to cell (Ctrl+G for addresses like A100, 10,5)
- Lazy loading for large files (1000+ rows)
- Help screen (? key)
- Export to CSV, JSON, text (non-interactive mode)

## Future Enhancements

### Export from TUI
- Add export menu accessible from TUI mode (`e` key)
- Allow choosing format (CSV, JSON, text) interactively
- Support file path selection
- Show progress for large exports

### Formula Toggle Mode
- Add `f` key to toggle between values and formulas in grid
- Currently formulas only visible in cell detail popup (Enter key)
- Full grid formula view would match Excel's formula bar behavior

### Visual Polish
- Color themes (light/dark toggle)
- Cell type coloring (numbers, strings, dates, errors)
- Grid lines toggle
- Improved header styling
- Alternating row colors

### Advanced Navigation & Filtering
- Filter mode to show only matching rows
- Sort by column (click header or hotkey)
- Freeze panes (keep headers visible while scrolling)

### Configuration Support
- `~/.config/xleak/config.toml` for user preferences
- Customizable keybindings
- Default theme selection
- Default column width settings

### Mouse Support
- Click cells to select
- Mouse wheel scrolling
- Click sheet tabs to switch
- Double-click for cell detail view

### Testing Expansion
- More comprehensive unit tests
- Integration tests for TUI interactions
- Edge case testing (empty sheets, single cells, huge files)
- Cross-platform testing automation

## Non-Goals

- Editing Excel files (read-only tool)
- Complex formula evaluation (just display them)
- Chart/graph rendering (terminal limitations)
- Macro execution (security risk)

## Contributing

See AGENTS.md for development guidelines and architecture notes.
