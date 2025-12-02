# Multi-line Cell Display Enhancement

**Status:** Planned (not implemented)
**Related Issue:** #16
**Estimated Complexity:** Medium (2-3 hours)
**Priority:** Low (enhancement, not critical)

## Overview

Add optional flag to expand multi-line cells in the main TUI table view, allowing users to see multiple lines per cell instead of just the first line.

## Current Behavior

**Main Table View:**
- Multi-line cells show only the first line (or newlines render as spaces/weird characters)
- Row height is fixed at 1 line
- Users must press Enter to open cell detail popup to see full content

**Cell Detail Popup:**
- ✅ **Fixed in Issue #16**: Now scrollable to view all lines in multi-line cells
- Supports Up/Down/PgUp/PgDn/Home keys
- Shows scroll position indicator

## Proposed Feature

### CLI Flag

```bash
--max-cell-lines N
```

**Behavior:**
- `N = 1` (default): Current behavior, single-line cells
- `N = 2-99`: Show up to N lines per cell in main table view
- `N = 0`: Unlimited (show all lines)

**Examples:**
```bash
# Default: single line per cell
xleak data.xlsx -i

# Show up to 3 lines per cell
xleak data.xlsx -i --max-cell-lines 3

# Show all lines (unlimited)
xleak data.xlsx -i --max-cell-lines 0
```

### Scope

**TUI Mode Only** (initially):
- Interactive TUI would benefit most from this feature
- Non-interactive mode (`display.rs`) uses prettytable-rs which has limited multi-line support
- Can extend to non-interactive mode later if needed

## Implementation Details

### Files to Modify

#### 1. `src/main.rs` (~10 lines)
Add CLI argument:
```rust
/// Maximum lines to show per cell in TUI (0 = unlimited, default: 1)
#[arg(long, default_value = "1")]
max_cell_lines: usize,
```

Pass to TUI:
```rust
tui::run_tui(wb, &sheet_name, &config, cli.horizontal_scroll, cli.max_cell_lines)?;
```

#### 2. `src/tui.rs` (~60 lines total)

**Add state field** (around line 471):
```rust
pub struct TuiState {
    // ... existing fields ...
    max_cell_lines: usize,
}
```

**Update constructor** (line 508):
```rust
pub fn new(
    mut workbook: Workbook,
    initial_sheet_name: &str,
    config: &crate::config::Config,
    horizontal_scroll: bool,
    max_cell_lines: usize, // Add parameter
) -> Result<Self>
```

**Update run_tui signature** (line 2079):
```rust
pub fn run_tui(
    workbook: Workbook,
    sheet_name: &str,
    config: &crate::config::Config,
    horizontal_scroll: bool,
    max_cell_lines: usize, // Add parameter
) -> Result<()>
```

**Modify row rendering** (lines 1315-1374):
The main implementation - calculate dynamic row heights:
```rust
let data_rows: Vec<Row> = visible_rows
    .iter()
    .enumerate()
    .map(|(visible_idx, row)| {
        let row_idx = visible_start + visible_idx;

        // Calculate row height based on max lines in any cell
        let mut row_height = 1;
        if self.max_cell_lines != 1 {
            for cell in row.iter() {
                let cell_str = cell.to_string();
                let line_count = if self.max_cell_lines == 0 {
                    cell_str.lines().count() // Unlimited
                } else {
                    cell_str.lines().count().min(self.max_cell_lines)
                };
                row_height = row_height.max(line_count);
            }
        }

        let cells: Vec<Cell> = row
            .iter()
            .enumerate()
            .map(|(col_idx, cell)| {
                let cell_str = cell.to_string();

                // Handle multi-line cells
                let display_text = if self.max_cell_lines == 1 {
                    // Single line: replace newlines with space
                    cell_str.replace('\n', " ")
                } else {
                    // Multi-line: take first N lines
                    let lines: Vec<&str> = cell_str.lines().collect();
                    let take_lines = if self.max_cell_lines == 0 {
                        lines.len()
                    } else {
                        lines.len().min(self.max_cell_lines)
                    };
                    lines[..take_lines].join("\n")
                };

                // ... existing styling logic ...
                Cell::from(display_text).style(style)
            })
            .collect();

        Row::new(cells).height(row_height as u16) // Use calculated height
    })
    .collect();
```

**Update column width calculation** (line 960):
Handle multi-line content when calculating widths:
```rust
let len = cell.to_string()
    .lines()
    .map(|line| line.len())
    .max()
    .unwrap_or(0);
```

## Technical Considerations

### Performance
- Only calculates line counts when `max_cell_lines > 1` (zero overhead by default)
- For visible cells only (not entire dataset)
- Minimal impact on render performance

### UX Impact
- **With large N or unlimited**: Rows could become very tall (20+ lines)
- **Scrolling**: Works fine with variable row heights (viewport shows fewer complete rows)
- **Alignment**: ratatui handles multi-line cells automatically (top-aligned)

### Edge Cases to Handle
- Empty cells
- Cells with only newlines
- Mixed content (some cells single-line, some multi-line in same row)
- Column width calculation with very long lines

## Testing Plan

1. **Manual testing** with `tests/fixtures/multiline_test.xlsx`:
   - Default behavior (--max-cell-lines 1)
   - Limited lines (--max-cell-lines 3)
   - Unlimited (--max-cell-lines 0)
   - With horizontal scroll enabled
   - With large files (performance check)

2. **Edge cases:**
   - Empty cells in multi-line rows
   - Single newline characters
   - Very long lines (horizontal scroll interaction)

3. **Verify existing tests still pass**

## Documentation Updates

### README.md
Add to TUI mode usage section:
```markdown
# Show up to 5 lines per cell in table view
xleak data.xlsx -i --max-cell-lines 5

# Show all lines (unlimited)
xleak data.xlsx -i --max-cell-lines 0
```

### Help screen (src/tui.rs)
Add note about multi-line display if flag is enabled

## Future Enhancements

- Add to non-interactive table display (prettytable-rs or custom renderer)
- Visual indicator in single-line mode showing "..." or "↕" for multi-line cells
- Configuration file option (default max-cell-lines value)
- Smart truncation (show first N-1 lines + "... (X more lines)")

## References

- **Issue #16:** https://github.com/bgreenwell/xleak/issues/16
- **Reporter:** @ket000
- **Current fix:** Cell detail popup scrolling (implemented)
- **This enhancement:** Optional main table view multi-line support (planned)
