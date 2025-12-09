use anyhow::{Context, Result, anyhow};
use calamine::{Data, Range, Reader, Sheets, Table, open_workbook_auto};
use chrono::{Duration, NaiveDate};
use std::path::Path;

/// Attempts to parse a string into a numeric CellValue, otherwise returns it as a String.
fn parse_string_to_cellvalue(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }
    // Try parsing as an integer first
    if let Ok(i) = s.parse::<i64>() {
        return CellValue::Int(i);
    }
    // Then try as a float
    if let Ok(f) = s.parse::<f64>() {
        return CellValue::Float(f);
    }
    // Default to a string
    CellValue::String(s.to_string())
}

/// Loads a CSV file into a CsvData object.
fn load_csv_data(path: &Path) -> Result<CsvData> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(true)
        .from_path(path)?;

    let headers = reader
        .headers()?
        .iter()
        .map(String::from)
        .collect::<Vec<String>>();
    let width = headers.len();

    let mut rows = Vec::new();
    for result in reader.records() {
        let record = result?;
        let row: Vec<CellValue> = record.iter().map(parse_string_to_cellvalue).collect();
        rows.push(row);
    }

    let height = rows.len();

    let sheet_data = SheetData {
        headers,
        rows,
        formulas: vec![vec![None; width]; height], // CSVs don't have formulas
        width,
        height,
    };

    let name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("data")
        .to_string();

    Ok(CsvData {
        name,
        data: sheet_data,
    })
}

// +++++ Refactored Workbook and Data Structures +++++

#[derive(Debug, Clone)]
pub struct CsvData {
    pub name: String,
    pub data: SheetData,
}

pub enum DataSource {
    Excel(Sheets<std::io::BufReader<std::fs::File>>),
    Csv(CsvData),
}

pub struct Workbook {
    pub source: DataSource,
}

impl Workbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let source = if path.extension().and_then(|s| s.to_str()) == Some("csv") {
            let csv_data = load_csv_data(path).with_context(|| "Failed to load CSV file")?;
            DataSource::Csv(csv_data)
        } else {
            let sheets = open_workbook_auto(path).context("Failed to open workbook")?;
            DataSource::Excel(sheets)
        };

        Ok(Self { source })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        match &self.source {
            DataSource::Excel(sheets) => sheets.sheet_names(),
            DataSource::Csv(csv_data) => vec![csv_data.name.clone()],
        }
    }

    /// Loads all rows eagerly into memory
    pub fn load_sheet(&mut self, name: &str) -> Result<SheetData> {
        match &mut self.source {
            DataSource::Excel(sheets) => {
                let range = sheets
                    .worksheet_range(name)
                    .with_context(|| format!("Sheet '{name}' not found"))?;
                let formula_range = sheets.worksheet_formula(name).ok();
                Ok(SheetData::from_range_with_formulas(range, formula_range))
            }
            DataSource::Csv(csv_data) => {
                if csv_data.name == name {
                    Ok(csv_data.data.clone())
                } else {
                    Err(anyhow!("Sheet '{name}' not found in CSV."))
                }
            }
        }
    }

    /// Loads only headers; rows fetched on demand
    pub fn load_sheet_lazy(&mut self, name: &str) -> Result<LazySheetData> {
        match &mut self.source {
            DataSource::Excel(sheets) => {
                let range = sheets
                    .worksheet_range(name)
                    .with_context(|| format!("Sheet '{name}' not found"))?;
                let formula_range = sheets.worksheet_formula(name).ok();
                Ok(LazySheetData::from_excel(range, formula_range))
            }
            DataSource::Csv(csv_data) => {
                if csv_data.name == name {
                    Ok(LazySheetData::from_csv(csv_data.data.clone()))
                } else {
                    Err(anyhow!("Sheet '{name}' not found in CSV."))
                }
            }
        }
    }

    // ===== Table API (Xlsx only) =====

    pub fn load_tables(&mut self) -> Result<()> {
        match &mut self.source {
            DataSource::Excel(Sheets::Xlsx(xlsx)) => xlsx
                .load_tables()
                .context("Failed to load table metadata")
                .map_err(|e| anyhow!("{e}")),
            _ => Err(anyhow!("Tables are only supported in .xlsx files")),
        }
    }

    pub fn table_names(&self) -> Result<Vec<String>> {
        match &self.source {
            DataSource::Excel(Sheets::Xlsx(xlsx)) => {
                Ok(xlsx.table_names().iter().map(|s| (*s).clone()).collect())
            }
            _ => Err(anyhow!("Tables are only supported in .xlsx files")),
        }
    }

    pub fn table_names_in_sheet(&self, sheet_name: &str) -> Result<Vec<String>> {
        match &self.source {
            DataSource::Excel(Sheets::Xlsx(xlsx)) => Ok(xlsx
                .table_names_in_sheet(sheet_name)
                .iter()
                .map(|s| (*s).clone())
                .collect()),
            _ => Err(anyhow!("Tables are only supported in .xlsx files")),
        }
    }

    pub fn table_by_name(&mut self, table_name: &str) -> Result<TableData> {
        match &mut self.source {
            DataSource::Excel(Sheets::Xlsx(xlsx)) => {
                let table = xlsx
                    .table_by_name(table_name)
                    .map_err(|e| anyhow!("Table '{table_name}' not found: {e}"))?;
                Ok(TableData::from_calamine_table(table))
            }
            _ => Err(anyhow!("Tables are only supported in .xlsx files")),
        }
    }
}

/// Eagerly-loaded sheet data (loads all rows immediately)
#[derive(Debug, Clone)]
pub struct SheetData {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<CellValue>>,
    pub formulas: Vec<Vec<Option<String>>>, // Parallel structure to rows with formulas
    pub width: usize,
    pub height: usize,
}

enum LazyDataSource {
    Excel {
        range: Range<Data>,
        formula_range: Option<Range<String>>,
    },
    Csv {
        data: SheetData,
    },
}

/// Lazy-loaded sheet data (loads rows on demand)
pub struct LazySheetData {
    source: LazyDataSource,
    pub headers: Vec<String>,
    pub width: usize,
    pub height: usize,
}

impl LazySheetData {
    /// Create lazy data from an Excel range
    pub fn from_excel(range: Range<Data>, formula_range: Option<Range<String>>) -> Self {
        let (height, width) = range.get_size();
        let headers = if height > 0 {
            range
                .rows()
                .next()
                .map(|row| row.iter().map(SheetData::cell_to_string).collect())
                .unwrap_or_default()
        } else {
            vec![]
        };

        Self {
            source: LazyDataSource::Excel {
                range,
                formula_range,
            },
            headers,
            width,
            height: height.saturating_sub(1),
        }
    }

    /// Create "lazy" data from already-loaded CSV data
    pub fn from_csv(data: SheetData) -> Self {
        Self {
            headers: data.headers.clone(),
            width: data.width,
            height: data.height,
            source: LazyDataSource::Csv { data },
        }
    }

    /// Zero-indexed row range; header excluded
    pub fn get_rows(
        &self,
        start: usize,
        count: usize,
    ) -> (Vec<Vec<CellValue>>, Vec<Vec<Option<String>>>) {
        match &self.source {
            LazyDataSource::Excel {
                range,
                formula_range,
            } => self.get_excel_rows(start, count, range, formula_range),
            LazyDataSource::Csv { data } => self.get_csv_rows(start, count, data),
        }
    }

    fn get_csv_rows(
        &self,
        start: usize,
        count: usize,
        data: &SheetData,
    ) -> (Vec<Vec<CellValue>>, Vec<Vec<Option<String>>>) {
        let end = (start + count).min(self.height);
        let rows = data.rows[start..end].to_vec();
        let formulas = data.formulas[start..end].to_vec();
        (rows, formulas)
    }

    fn get_excel_rows(
        &self,
        start: usize,
        count: usize,
        range: &Range<Data>,
        formula_range: &Option<Range<String>>,
    ) -> (Vec<Vec<CellValue>>, Vec<Vec<Option<String>>>) {
        let end = (start + count).min(self.height);

        let rows: Vec<Vec<CellValue>> = range
            .rows()
            .skip(1 + start)
            .take(end - start)
            .map(|row| row.iter().map(SheetData::datatype_to_cellvalue).collect())
            .collect();

        let formulas = self.get_formulas_for_range(start, end, formula_range);

        (rows, formulas)
    }

    fn get_formulas_for_range(
        &self,
        start: usize,
        end: usize,
        formula_range: &Option<Range<String>>,
    ) -> Vec<Vec<Option<String>>> {
        if let Some(formula_range) = formula_range {
            let formula_start = formula_range.start().unwrap_or((0, 0));
            let total_height = self.height + 1;

            let mut formula_grid: Vec<Vec<Option<String>>> =
                vec![vec![None; self.width]; end - start];

            for (row_offset, formula_row) in formula_range.rows().enumerate() {
                let absolute_row = formula_start.0 as usize + row_offset;

                if absolute_row > 0 && absolute_row <= total_height {
                    let data_row_idx = absolute_row - 1;

                    if data_row_idx >= start && data_row_idx < end {
                        let result_idx = data_row_idx - start;

                        for (col_offset, formula_str) in formula_row.iter().enumerate() {
                            let absolute_col = formula_start.1 as usize + col_offset;
                            if absolute_col < self.width && !formula_str.is_empty() {
                                formula_grid[result_idx][absolute_col] = Some(formula_str.clone());
                            }
                        }
                    }
                }
            }

            formula_grid
        } else {
            vec![vec![None; self.width]; end - start]
        }
    }

    /// Consumes lazy data and loads all rows into memory
    #[allow(clippy::wrong_self_convention)]
    pub fn to_sheet_data(self) -> SheetData {
        match self.source {
            LazyDataSource::Excel {
                range,
                formula_range,
            } => SheetData::from_range_with_formulas(range, formula_range),
            LazyDataSource::Csv { data } => data,
        }
    }
}

#[derive(Debug, Clone)]
pub enum CellValue {
    Empty,
    String(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    Error(String),
    DateTime(f64), // Excel datetime as float
}

impl CellValue {
    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    #[allow(dead_code)]
    pub fn is_numeric(&self) -> bool {
        matches!(self, CellValue::Int(_) | CellValue::Float(_))
    }

    /// Returns unformatted value (for export/clipboard)
    pub fn to_raw_string(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Int(i) => i.to_string(),
            CellValue::Float(val) => {
                if val.fract() == 0.0 {
                    format!("{val:.0}")
                } else {
                    val.to_string()
                }
            }
            CellValue::Bool(b) => b.to_string(),
            CellValue::Error(e) => format!("#{e}"),
            CellValue::DateTime(dt) => {
                let days = dt.floor() as i64;
                let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
                let adjusted_days = if days > 60 { days - 1 } else { days };
                let date = epoch + Duration::days(adjusted_days);
                let time_fraction = dt.fract();
                let total_seconds = (time_fraction * 86400.0).round() as i64;
                let hours = total_seconds / 3600;
                let minutes = (total_seconds % 3600) / 60;
                let seconds = total_seconds % 60;

                if time_fraction.abs() < 0.0000001 {
                    format!("{}", date.format("%Y-%m-%d"))
                } else {
                    format!(
                        "{} {:02}:{:02}:{:02}",
                        date.format("%Y-%m-%d"),
                        hours,
                        minutes,
                        seconds
                    )
                }
            }
        }
    }
}

/// Excel Table data
#[derive(Debug, Clone)]
pub struct TableData {
    pub name: String,
    pub sheet_name: String,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<CellValue>>,
}

impl TableData {
    pub fn from_calamine_table(table: Table<Data>) -> Self {
        let name = table.name().to_string();
        let sheet_name = table.sheet_name().to_string();
        let headers = table.columns().to_vec();

        let rows: Vec<Vec<CellValue>> = table
            .data()
            .rows()
            .map(|row| row.iter().map(SheetData::datatype_to_cellvalue).collect())
            .collect();

        Self {
            name,
            sheet_name,
            headers,
            rows,
        }
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::String(s) => write!(f, "{s}"),
            CellValue::Int(i) => {
                let s = i.to_string();
                let negative = s.starts_with('-');
                let digits: String = s.trim_start_matches('-').chars().collect();
                let mut result = String::new();
                for (idx, ch) in digits.chars().rev().enumerate() {
                    if idx > 0 && idx % 3 == 0 {
                        result.push(',');
                    }
                    result.push(ch);
                }
                if negative {
                    result.push('-');
                }
                write!(f, "{}", result.chars().rev().collect::<String>())
            }
            CellValue::Float(val) => {
                let formatted = if val.fract() == 0.0 {
                    format!("{val:.0}")
                } else {
                    format!("{val:.2}")
                };
                let parts: Vec<&str> = formatted.split('.').collect();
                let int_part = parts[0];
                let negative = int_part.starts_with('-');
                let digits: String = int_part.trim_start_matches('-').chars().collect();
                let mut result = String::new();
                for (idx, ch) in digits.chars().rev().enumerate() {
                    if idx > 0 && idx % 3 == 0 {
                        result.push(',');
                    }
                    result.push(ch);
                }
                if negative {
                    result.push('-');
                }
                let int_formatted: String = result.chars().rev().collect();
                if parts.len() > 1 {
                    write!(f, "{}.{}", int_formatted, parts[1])
                } else {
                    write!(f, "{}", int_formatted)
                }
            }
            CellValue::Bool(b) => {
                write!(f, "{}", if *b { "true" } else { "false" })
            }
            CellValue::Error(e) => write!(f, "ERROR: {e}"),
            CellValue::DateTime(d) => {
                let days = d.floor() as i64;
                let excel_epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
                let adjusted_days = if days > 60 { days - 1 } else { days };

                if let Some(date) = excel_epoch.checked_add_signed(Duration::days(adjusted_days)) {
                    let frac = d.fract();
                    if frac.abs() > 0.000001 {
                        let total_seconds = (frac * 86400.0).round() as u32;
                        let hours = total_seconds / 3600;
                        let minutes = (total_seconds % 3600) / 60;
                        let seconds = total_seconds % 60;
                        write!(f, "{} {:02}:{:02}:{:02}", date, hours, minutes, seconds)
                    } else {
                        write!(f, "{}", date)
                    }
                } else {
                    write!(f, "Date[{days}]")
                }
            }
        }
    }
}

impl SheetData {
    pub fn from_range_with_formulas(
        range: Range<Data>,
        formula_range: Option<Range<String>>,
    ) -> Self {
        let (height, width) = range.get_size();

        let headers = if height > 0 {
            range
                .rows()
                .next()
                .map(|row| row.iter().map(Self::cell_to_string).collect())
                .unwrap_or_default()
        } else {
            vec![]
        };

        let rows: Vec<Vec<CellValue>> = range
            .rows()
            .skip(1)
            .map(|row| row.iter().map(Self::datatype_to_cellvalue).collect())
            .collect();

        let formulas: Vec<Vec<Option<String>>> = if let Some(formula_range) = formula_range {
            let formula_start = formula_range.start().unwrap_or((0, 0));
            let mut formula_grid: Vec<Vec<Option<String>>> = vec![vec![None; width]; height];

            for (row_offset, formula_row) in formula_range.rows().enumerate() {
                let absolute_row = formula_start.0 as usize + row_offset;
                if absolute_row > 0 && absolute_row <= height {
                    let data_row_idx = absolute_row - 1;
                    for (col_offset, formula_str) in formula_row.iter().enumerate() {
                        let absolute_col = formula_start.1 as usize + col_offset;
                        if absolute_col < width && !formula_str.is_empty() {
                            formula_grid[data_row_idx][absolute_col] = Some(formula_str.clone());
                        }
                    }
                }
            }

            formula_grid
                .into_iter()
                .take(height.saturating_sub(1))
                .collect()
        } else {
            vec![vec![None; width]; height.saturating_sub(1)]
        };

        Self {
            headers,
            rows,
            formulas,
            width,
            height: height.saturating_sub(1),
        }
    }

    fn cell_to_string(cell: &Data) -> String {
        match cell {
            Data::Empty => String::new(),
            Data::String(s) => s.clone(),
            Data::Int(i) => i.to_string(),
            Data::Float(f) => {
                if f.fract() == 0.0 {
                    format!("{f:.0}")
                } else {
                    f.to_string()
                }
            }
            Data::Bool(b) => b.to_string(),
            Data::Error(e) => format!("ERROR: {e:?}"),
            Data::DateTime(d) => format!("Date({})", d.as_f64()),
            Data::DateTimeIso(s) => s.clone(),
            Data::DurationIso(s) => s.clone(),
        }
    }

    fn datatype_to_cellvalue(cell: &Data) -> CellValue {
        match cell {
            Data::Empty => CellValue::Empty,
            Data::String(s) => CellValue::String(s.clone()),
            Data::Int(i) => CellValue::Int(*i),
            Data::Float(f) => CellValue::Float(*f),
            Data::Bool(b) => CellValue::Bool(*b),
            Data::Error(e) => CellValue::Error(format!("{e:?}")),
            Data::DateTime(d) => CellValue::DateTime(d.as_f64()),
            Data::DateTimeIso(s) => CellValue::String(s.clone()),
            Data::DurationIso(s) => CellValue::String(s.clone()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cellvalue_display_integer() {
        let val = CellValue::Int(1234567);
        assert_eq!(val.to_string(), "1,234,567");
    }

    #[test]
    fn test_cellvalue_display_negative_integer() {
        let val = CellValue::Int(-1234567);
        assert_eq!(val.to_string(), "-1,234,567");
    }

    #[test]
    fn test_cellvalue_display_float() {
        let val = CellValue::Float(1234567.89);
        assert_eq!(val.to_string(), "1,234,567.89");
    }

    #[test]
    fn test_cellvalue_display_float_whole_number() {
        let val = CellValue::Float(1000.0);
        assert_eq!(val.to_string(), "1,000");
    }

    #[test]
    fn test_cellvalue_display_boolean() {
        assert_eq!(CellValue::Bool(true).to_string(), "true");
        assert_eq!(CellValue::Bool(false).to_string(), "false");
    }

    #[test]
    fn test_cellvalue_display_string() {
        let val = CellValue::String("Hello, World!".to_string());
        assert_eq!(val.to_string(), "Hello, World!");
    }

    #[test]
    fn test_cellvalue_display_empty() {
        let val = CellValue::Empty;
        assert_eq!(val.to_string(), "");
    }

    #[test]
    fn test_cellvalue_display_error() {
        let val = CellValue::Error("DIV/0!".to_string());
        assert_eq!(val.to_string(), "ERROR: DIV/0!");
    }

    #[test]
    fn test_cellvalue_to_raw_string_integer() {
        let val = CellValue::Int(1234567);
        assert_eq!(val.to_raw_string(), "1234567");
    }

    #[test]
    fn test_cellvalue_to_raw_string_float() {
        let val = CellValue::Float(123.45);
        assert_eq!(val.to_raw_string(), "123.45");
    }

    #[test]
    fn test_cellvalue_is_empty() {
        assert!(CellValue::Empty.is_empty());
        assert!(!CellValue::Int(0).is_empty());
        assert!(!CellValue::String("".to_string()).is_empty());
    }

    #[test]
    fn test_cellvalue_is_numeric() {
        assert!(CellValue::Int(123).is_numeric());
        assert!(CellValue::Float(123.45).is_numeric());
        assert!(!CellValue::String("123".to_string()).is_numeric());
        assert!(!CellValue::Empty.is_numeric());
    }

    #[test]
    fn test_datetime_display() {
        let val = CellValue::DateTime(1.0);
        let display = val.to_string();
        assert!(display.contains("1900") || display.contains("1899"));
    }

    #[test]
    fn test_datetime_with_time() {
        let val = CellValue::DateTime(1.5);
        let display = val.to_string();
        assert!(display.contains(":"));
        assert!(display.len() > 10);
    }

    #[test]
    fn test_workbook_open_real_file() {
        if let Ok(wb) = Workbook::open("tests/fixtures/test_data.xlsx") {
            let sheet_names = wb.sheet_names();
            assert!(!sheet_names.is_empty(), "Should have at least one sheet");
        }
    }

    #[test]
    fn test_sheet_data_structure() {
        let sheet = SheetData {
            headers: vec!["Name".to_string(), "Age".to_string()],
            rows: vec![
                vec![CellValue::String("Alice".to_string()), CellValue::Int(30)],
                vec![CellValue::String("Bob".to_string()), CellValue::Int(25)],
            ],
            formulas: vec![vec![None, None], vec![None, None]],
            width: 2,
            height: 2,
        };

        assert_eq!(sheet.width, 2);
        assert_eq!(sheet.height, 2);
        assert_eq!(sheet.headers.len(), 2);
        assert_eq!(sheet.rows.len(), 2);
    }
}
