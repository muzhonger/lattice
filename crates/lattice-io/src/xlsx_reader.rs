//! Read `.xlsx` and `.ods` files using calamine and convert to Lattice `Workbook`.
//!
//! Cell values come from calamine (`calamine::Reader`). Three things are then
//! recovered by extra passes over the raw XML, because calamine does not expose
//! them: formula *text* (calamine only returns computed values), and cell
//! styles / sheet layout (merged regions, column widths, row heights, hidden
//! rows and columns, tab colour).
//!
//! Deliberately **not** imported (kept honest — see the fidelity matrix in
//! `docs/table/`): comments, hyperlinks, conditional formatting, data
//! validation, charts, images, pivot tables, print settings, colour *tints* on
//! theme colours, and styles on cells that have no value (attaching a format to
//! an empty cell would inflate `used_range`).

use std::collections::{HashMap, HashSet};
use std::io::Read;
#[cfg(feature = "native")]
use std::path::Path;

use calamine::{Data, Reader};
use quick_xml::Reader as XmlReader;
use quick_xml::events::{BytesStart, Event};

use lattice_core::{
    Border, BorderStyle, Cell, CellBorders, CellError, CellFormat, CellValue, HAlign, Sheet,
    TextWrap, VAlign, Workbook,
};

use crate::{IoError, Result};

/// Read an `.xlsx` file and return a populated `Workbook`.
///
/// Each sheet in the Excel file becomes a sheet in the workbook.
/// Cell values are converted from calamine's `Data` enum to our `CellValue`.
#[cfg(feature = "native")]
pub fn read_xlsx(path: &Path) -> Result<Workbook> {
    if !path.exists() {
        return Err(IoError::FileNotFound(path.display().to_string()));
    }

    let bytes = std::fs::read(path)?;
    read_xlsx_from_bytes(&bytes)
}

/// Read an `.xlsx` file from in-memory bytes and return a populated `Workbook`.
///
/// This is the shared core used by [`read_xlsx`] (which reads the file first)
/// and is also the entry point for the WASM build, which has no filesystem.
/// calamine reads any `Read + Seek` source, so a `Cursor` over the bytes works.
pub fn read_xlsx_from_bytes(bytes: &[u8]) -> Result<Workbook> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut excel: calamine::Xlsx<_> = calamine::Xlsx::new(cursor)
        .map_err(|e: calamine::XlsxError| IoError::XlsxRead(e.to_string()))?;

    let sheet_names = excel.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err(IoError::XlsxRead("workbook has no sheets".into()));
    }

    let mut workbook = Workbook::new();

    // Add all sheets from the file.
    for (i, name) in sheet_names.iter().enumerate() {
        if i == 0 {
            // Rename the default "Sheet1" to the first sheet name.
            if name != "Sheet1" {
                workbook
                    .rename_sheet("Sheet1", name.as_str())
                    .map_err(IoError::Core)?;
            }
        } else {
            workbook.add_sheet(name.as_str()).map_err(IoError::Core)?;
        }
    }

    // Populate each sheet with data.
    for name in &sheet_names {
        let range: calamine::Range<Data> = match excel.worksheet_range(name) {
            Ok(r) => r,
            Err(e) => {
                // Skip sheets that can't be read (e.g. chart sheets).
                eprintln!("warning: skipping sheet '{}': {}", name, e);
                continue;
            }
        };

        let sheet = workbook.get_sheet_mut(name).map_err(IoError::Core)?;

        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, cell_data) in row.iter().enumerate() {
                let value = calamine_data_to_cell_value(cell_data);
                if value != CellValue::Empty {
                    let cell = Cell {
                        value,
                        ..Default::default()
                    };
                    sheet.set_cell(row_idx as u32, col_idx as u32, cell);
                }
            }
        }
    }

    // Set active sheet to the first one.
    workbook.active_sheet = sheet_names[0].clone();

    // Post-process: extract formula text from the raw xlsx XML.
    // calamine only returns computed values, not the formula strings.
    if let Err(e) = extract_formulas_from_bytes(bytes, &mut workbook) {
        eprintln!("warning: could not extract formulas: {}", e);
    }

    // Post-process: cell styles + sheet layout (merges, sizes, hidden, tab colour).
    // Applied after formulas so formula-only cells also receive their format.
    if let Err(e) = apply_styles_and_layout_from_bytes(bytes, &mut workbook) {
        eprintln!("warning: could not read styles/layout: {}", e);
    }

    Ok(workbook)
}

/// Read an `.ods` (OpenDocument Spreadsheet) file and return a populated `Workbook`.
///
/// Uses calamine's ODS support. Cell values are converted the same way as xlsx.
#[cfg(feature = "native")]
pub fn read_ods(path: &Path) -> Result<Workbook> {
    if !path.exists() {
        return Err(IoError::FileNotFound(path.display().to_string()));
    }

    let bytes = std::fs::read(path)?;
    read_ods_from_bytes(&bytes)
}

/// Read an `.ods` (OpenDocument Spreadsheet) file from in-memory bytes.
///
/// WASM-available counterpart of [`read_ods`].
pub fn read_ods_from_bytes(bytes: &[u8]) -> Result<Workbook> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut ods: calamine::Ods<_> = calamine::Ods::new(cursor)
        .map_err(|e: calamine::OdsError| IoError::XlsxRead(format!("ODS error: {}", e)))?;

    let sheet_names = ods.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err(IoError::XlsxRead("ODS workbook has no sheets".into()));
    }

    let mut workbook = Workbook::new();

    for (i, name) in sheet_names.iter().enumerate() {
        if i == 0 {
            if name != "Sheet1" {
                workbook
                    .rename_sheet("Sheet1", name.as_str())
                    .map_err(IoError::Core)?;
            }
        } else {
            workbook.add_sheet(name.as_str()).map_err(IoError::Core)?;
        }
    }

    for name in &sheet_names {
        let range: calamine::Range<Data> = match ods.worksheet_range(name) {
            Ok(r) => r,
            Err(e) => {
                eprintln!("warning: skipping ODS sheet '{}': {}", name, e);
                continue;
            }
        };

        let sheet = workbook.get_sheet_mut(name).map_err(IoError::Core)?;

        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, cell_data) in row.iter().enumerate() {
                let value = calamine_data_to_cell_value(cell_data);
                if value != CellValue::Empty {
                    let cell = Cell {
                        value,
                        ..Default::default()
                    };
                    sheet.set_cell(row_idx as u32, col_idx as u32, cell);
                }
            }
        }
    }

    workbook.active_sheet = sheet_names[0].clone();
    Ok(workbook)
}

/// Auto-detect format (xlsx, xls, ods) and read the file.
///
/// Uses [`crate::format_detect::detect_format`] to pick the right reader.
#[cfg(feature = "native")]
pub fn read_spreadsheet(path: &Path) -> Result<Workbook> {
    use crate::format_detect::{FileFormat, detect_format};

    let format = detect_format(path)?;
    match format {
        FileFormat::Xlsx => read_xlsx(path),
        FileFormat::Xls => read_xls(path),
        FileFormat::Ods => read_ods(path),
        FileFormat::Csv => crate::csv_io::read_csv(path),
        FileFormat::Tsv => crate::tsv_io::read_tsv(path),
        FileFormat::Json => Err(IoError::UnsupportedFormat(
            "JSON import is not supported; use CSV or XLSX".to_string(),
        )),
    }
}

/// Read a legacy `.xls` file using calamine's XLS support.
#[cfg(feature = "native")]
pub fn read_xls(path: &Path) -> Result<Workbook> {
    if !path.exists() {
        return Err(IoError::FileNotFound(path.display().to_string()));
    }

    let bytes = std::fs::read(path)?;
    read_xls_from_bytes(&bytes)
}

/// Read a legacy `.xls` file from in-memory bytes.
///
/// WASM-available counterpart of [`read_xls`].
pub fn read_xls_from_bytes(bytes: &[u8]) -> Result<Workbook> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut xls: calamine::Xls<_> = calamine::Xls::new(cursor)
        .map_err(|e: calamine::XlsError| IoError::XlsxRead(format!("XLS error: {}", e)))?;

    let sheet_names = xls.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err(IoError::XlsxRead("XLS workbook has no sheets".into()));
    }

    let mut workbook = Workbook::new();

    for (i, name) in sheet_names.iter().enumerate() {
        if i == 0 {
            if name != "Sheet1" {
                workbook
                    .rename_sheet("Sheet1", name.as_str())
                    .map_err(IoError::Core)?;
            }
        } else {
            workbook.add_sheet(name.as_str()).map_err(IoError::Core)?;
        }
    }

    for name in &sheet_names {
        let range: calamine::Range<Data> = match xls.worksheet_range(name) {
            Ok(r) => r,
            Err(e) => {
                eprintln!("warning: skipping XLS sheet '{}': {}", name, e);
                continue;
            }
        };

        let sheet = workbook.get_sheet_mut(name).map_err(IoError::Core)?;

        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, cell_data) in row.iter().enumerate() {
                let value = calamine_data_to_cell_value(cell_data);
                if value != CellValue::Empty {
                    let cell = Cell {
                        value,
                        ..Default::default()
                    };
                    sheet.set_cell(row_idx as u32, col_idx as u32, cell);
                }
            }
        }
    }

    workbook.active_sheet = sheet_names[0].clone();
    Ok(workbook)
}

/// Extract formula text from the raw xlsx XML and set it on the workbook cells.
///
/// Opens the xlsx bytes as a ZIP archive, reads `xl/workbook.xml` to map sheet
/// names to relationship IDs, then reads `xl/_rels/workbook.xml.rels` to resolve
/// each rId to a worksheet XML path. Finally, parses each worksheet XML for
/// `<c>` elements containing `<f>` children and sets `cell.formula` accordingly.
fn extract_formulas_from_bytes(bytes: &[u8], workbook: &mut Workbook) -> Result<()> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|e| IoError::XlsxRead(format!("zip error: {}", e)))?;

    // Step 1: Parse xl/workbook.xml to build sheet name -> rId map.
    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml")?;
    let sheet_to_rid = parse_sheet_rid_map(&workbook_xml);

    // Step 2: Parse xl/_rels/workbook.xml.rels to build rId -> file path map.
    let rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rid_to_target = parse_rid_target_map(&rels_xml);

    // Step 3: For each sheet, find the worksheet XML and extract formulas.
    let sheet_names = workbook.sheet_names();
    for sheet_name in &sheet_names {
        let rid = match sheet_to_rid.get(sheet_name.as_str()) {
            Some(r) => r,
            None => continue,
        };
        let target = match rid_to_target.get(rid.as_str()) {
            Some(t) => t,
            None => continue,
        };

        // Target is relative to xl/, e.g. "worksheets/sheet1.xml"
        let xml_path = format!("xl/{}", target);
        let sheet_xml = match read_zip_entry_string(&mut archive, &xml_path) {
            Ok(xml) => xml,
            Err(_) => continue,
        };

        let formulas = parse_formulas_from_sheet_xml(&sheet_xml);
        if formulas.is_empty() {
            continue;
        }

        let sheet = match workbook.get_sheet_mut(sheet_name) {
            Ok(s) => s,
            Err(_) => continue,
        };

        for (cell_ref, formula_text) in &formulas {
            if let Some((row, col)) = parse_a1_ref(cell_ref) {
                // If the cell already exists (from calamine values), set formula on it.
                // If it doesn't exist, create a new cell with Empty value + formula.
                if let Some(cell) = sheet.get_cell_mut(row, col) {
                    cell.formula = Some(formula_text.clone());
                } else {
                    let cell = Cell {
                        formula: Some(formula_text.clone()),
                        ..Default::default()
                    };
                    sheet.set_cell(row, col, cell);
                }
            }
        }
    }

    Ok(())
}

/// Read a zip entry as a UTF-8 string.
fn read_zip_entry_string(
    archive: &mut zip::ZipArchive<std::io::Cursor<Vec<u8>>>,
    name: &str,
) -> Result<String> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| IoError::XlsxRead(format!("zip entry '{}': {}", name, e)))?;
    let mut buf = String::new();
    entry
        .read_to_string(&mut buf)
        .map_err(|e| IoError::XlsxRead(format!("reading '{}': {}", name, e)))?;
    Ok(buf)
}

/// Parse `xl/workbook.xml` to extract sheet name -> rId mapping.
///
/// Looks for `<sheet name="..." r:id="rIdN"/>` elements.
fn parse_sheet_rid_map(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let local = strip_ns(e.name().as_ref());
                if local == "sheet" {
                    let mut name = String::new();
                    let mut rid = String::new();
                    for attr in e.attributes().flatten() {
                        let key = strip_ns(attr.key.as_ref());
                        match key.as_str() {
                            "name" => {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            "id" => {
                                rid = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            _ => {}
                        }
                    }
                    if !name.is_empty() && !rid.is_empty() {
                        map.insert(name, rid);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Parse `xl/_rels/workbook.xml.rels` to extract rId -> target path mapping.
fn parse_rid_target_map(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    // Use extract_relationship_targets-style logic but capture Id -> Target.
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let local = strip_ns(e.name().as_ref());
                if local == "Relationship" {
                    let mut id = String::new();
                    let mut target = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                            b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                            _ => {}
                        }
                    }
                    if !id.is_empty() && !target.is_empty() {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Parse a worksheet XML string and extract cell formulas.
///
/// Returns a list of `(cell_ref, formula_text)` pairs, e.g.
/// `[("D5", "C5-B5"), ("E5", "D5/B5*100")]`.
///
/// The formula text does NOT include the leading `=`.
fn parse_formulas_from_sheet_xml(xml: &str) -> Vec<(String, String)> {
    let mut formulas = Vec::new();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut in_c = false;
    let mut current_ref = String::new();
    let mut in_f = false;
    let mut formula_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local = strip_ns(e.name().as_ref());
                match local.as_str() {
                    "c" => {
                        in_c = true;
                        current_ref.clear();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                current_ref = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    "f" if in_c => {
                        in_f = true;
                        formula_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_f && let Ok(text) = e.unescape() {
                    formula_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let local = strip_ns(e.name().as_ref());
                match local.as_str() {
                    "f" if in_f => {
                        in_f = false;
                        if !current_ref.is_empty() && !formula_text.is_empty() {
                            formulas.push((current_ref.clone(), formula_text.clone()));
                        }
                    }
                    "c" if in_c => {
                        in_c = false;
                        in_f = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    formulas
}

// ---------------------------------------------------------------------------
// Styles & layout import
// ---------------------------------------------------------------------------

/// Excel's legacy 64-entry `indexed="N"` palette.
const INDEXED_COLORS: [&str; 64] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
    "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
    "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
    "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
    "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
    "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
    "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333",
];

/// Theme slot names in `theme="N"` index order.
///
/// `<a:clrScheme>` lists its children as `dk1, lt1, dk2, lt2, accent1..6,
/// hlink, folHlink`, but the `theme` attribute indexes them as `lt1, dk1, lt2,
/// dk2, accent1..6, hlink, folHlink` — hence the remap by name.
const THEME_SLOTS: [&str; 12] = [
    "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
    "accent6", "hlink", "folHlink",
];

/// Office's default theme palette, used when the file has no `theme1.xml`.
const DEFAULT_THEME: [&str; 12] = [
    "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
    "70AD47", "0563C1", "954F72",
];

/// Cap on how many columns a single `<col min max>` element may expand to.
/// Guards against a whole-sheet `<col min="1" max="16384">` blowing up memory.
const MAX_COL_SPAN: u32 = 1024;

/// Normalise an Excel colour string (`AARRGGBB`, `RRGGBB`, `theme`-resolved) to
/// a CSS-style `#RRGGBB`.
fn normalize_hex(raw: &str) -> String {
    let hex: String = raw.chars().filter(|c| c.is_ascii_hexdigit()).collect();
    let tail = if hex.len() > 6 {
        &hex[hex.len() - 6..]
    } else {
        &hex[..]
    };
    format!("#{}", tail.to_ascii_uppercase())
}

/// Read an attribute by (namespace-stripped) name, unescaping XML entities.
fn attr_str(e: &BytesStart, key: &str) -> Option<String> {
    for a in e.attributes().flatten() {
        if strip_ns(a.key.as_ref()) == key {
            let raw = String::from_utf8_lossy(&a.value);
            return match quick_xml::escape::unescape(&raw) {
                Ok(v) => Some(v.into_owned()),
                Err(_) => Some(raw.to_string()),
            };
        }
    }
    None
}

fn attr_u32(e: &BytesStart, key: &str) -> Option<u32> {
    attr_str(e, key)?.trim().parse().ok()
}

fn attr_f64(e: &BytesStart, key: &str) -> Option<f64> {
    attr_str(e, key)?.trim().parse().ok()
}

/// A boolean OOXML attribute (`"1"`/`"true"` = true).
fn attr_truthy(e: &BytesStart, key: &str) -> bool {
    match attr_str(e, key) {
        Some(v) => !(v == "0" || v.eq_ignore_ascii_case("false")),
        None => false,
    }
}

/// A child element's `val` (e.g. `<b/>`, `<b val="0"/>`): absent `val` = on.
fn elem_enabled(e: &BytesStart) -> bool {
    match attr_str(e, "val") {
        Some(v) => !(v == "0" || v.eq_ignore_ascii_case("false")),
        None => true,
    }
}

/// Resolve a `<color …/>` element to `#RRGGBB`.
///
/// Supports `rgb` (the common case), `theme` (resolved through the theme
/// palette — colour *tints* are ignored) and the legacy `indexed` palette.
fn parse_color_element(e: &BytesStart, theme: &[String]) -> Option<String> {
    if let Some(rgb) = attr_str(e, "rgb") {
        return Some(normalize_hex(&rgb));
    }
    if let Some(idx) = attr_u32(e, "theme")
        && let Some(c) = theme.get(idx as usize)
    {
        return Some(normalize_hex(c));
    }
    if let Some(idx) = attr_u32(e, "indexed")
        && let Some(c) = INDEXED_COLORS.get(idx as usize)
    {
        return Some(normalize_hex(c));
    }
    if attr_truthy(e, "auto") {
        return Some("#000000".to_string());
    }
    None
}

fn border_style_from(name: &str) -> BorderStyle {
    match name {
        "thin" | "hair" => BorderStyle::Thin,
        "medium" | "mediumDashed" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "dotted" | "mediumDotted" => BorderStyle::Dotted,
        "dashed" | "dashDot" | "dashDotDot" | "slantDashDot" | "mediumDashDot"
        | "mediumDashDotDot" => BorderStyle::Dashed,
        _ => BorderStyle::None,
    }
}

/// Excel built-in number-format ids (ECMA-376 §18.8.30).
///
/// Custom formats (`numFmtId >= 164`) live in `<numFmts>` and take priority.
fn builtin_num_fmt(id: u32) -> Option<&'static str> {
    match id {
        0 => None, // General
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("m/d/yyyy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yyyy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

#[derive(Default)]
struct FontDef {
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    size: Option<f64>,
    name: Option<String>,
    color: Option<String>,
}

#[derive(Default)]
struct FillDef {
    solid: bool,
    fg: Option<String>,
}

#[derive(Default)]
struct AlignDef {
    h: Option<HAlign>,
    v: Option<VAlign>,
    wrap: bool,
    rotation: i16,
    indent: u8,
}

/// Per-`cellXfs`-entry attributes plus the optional `<alignment>` child.
#[derive(Default)]
struct XfDef {
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Option<AlignDef>,
}

/// Style dictionaries from `xl/styles.xml`, indexed the way `<c s="N">` does.
#[derive(Default)]
struct StyleDict {
    fonts: Vec<FontDef>,
    fills: Vec<Option<String>>,
    borders: Vec<CellBorders>,
    num_fmts: HashMap<u32, String>,
}

#[derive(Default, PartialEq, Clone, Copy)]
enum StyleSection {
    #[default]
    None,
    Fonts,
    Fills,
    Borders,
    NumFmts,
}

#[derive(Clone, Copy)]
enum BorderEdge {
    Left,
    Right,
    Top,
    Bottom,
}

#[derive(Default)]
struct DictState {
    section: StyleSection,
    font: Option<FontDef>,
    fill: Option<FillDef>,
    border: Option<CellBorders>,
    edge: Option<BorderEdge>,
}

/// Parse `xl/theme/theme1.xml` into the 12-colour palette in `theme="N"` order.
fn parse_theme_colors(xml: &str) -> Vec<String> {
    let mut slots: HashMap<String, String> = HashMap::new();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut in_scheme = false;
    let mut cur: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = strip_ns(e.name().as_ref());
                match name.as_str() {
                    "clrScheme" => in_scheme = true,
                    "srgbClr" | "sysClr" if in_scheme => {
                        let key = if name == "srgbClr" { "val" } else { "lastClr" };
                        if let (Some(slot), Some(v)) = (cur.as_ref(), attr_str(&e, key))
                            && !v.is_empty()
                        {
                            slots.insert(slot.clone(), v);
                        }
                    }
                    _ => {}
                }
                if in_scheme && THEME_SLOTS.contains(&name.as_str()) {
                    cur = Some(name);
                }
            }
            Ok(Event::Empty(e)) => {
                // `<a:sysClr …/>` / theme slots with no colour child.
                let name = strip_ns(e.name().as_ref());
                if in_scheme
                    && matches!(name.as_str(), "srgbClr" | "sysClr")
                    && let (Some(slot), Some(v)) = (
                        cur.as_ref(),
                        attr_str(&e, if name == "srgbClr" { "val" } else { "lastClr" }),
                    )
                    && !v.is_empty()
                {
                    slots.insert(slot.clone(), v);
                }
            }
            Ok(Event::End(e)) => {
                let name = strip_ns(e.name().as_ref());
                if name == "clrScheme" {
                    in_scheme = false;
                }
                if cur.as_deref() == Some(name.as_str()) {
                    cur = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    THEME_SLOTS
        .iter()
        .enumerate()
        .map(|(i, slot)| {
            slots
                .get(*slot)
                .map(|c| normalize_hex(c))
                .unwrap_or_else(|| normalize_hex(DEFAULT_THEME[i]))
        })
        .collect()
}

/// Pass 1 of `xl/styles.xml`: fonts, fills, borders and custom number formats.
fn parse_style_dict(xml: &str, theme: &[String]) -> StyleDict {
    let mut dict = StyleDict::default();
    let mut st = DictState::default();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => on_style_event(&mut dict, &mut st, &e, false, theme),
            Ok(Event::Empty(e)) => on_style_event(&mut dict, &mut st, &e, true, theme),
            Ok(Event::End(e)) => {
                let name = strip_ns(e.name().as_ref());
                match name.as_str() {
                    "fonts" | "fills" | "borders" | "numFmts" => st.section = StyleSection::None,
                    "font" if st.section == StyleSection::Fonts => {
                        dict.fonts.push(st.font.take().unwrap_or_default());
                    }
                    "fill" if st.section == StyleSection::Fills => {
                        let f = st.fill.take().unwrap_or_default();
                        dict.fills.push(if f.solid { f.fg } else { None });
                    }
                    "border" if st.section == StyleSection::Borders => {
                        dict.borders.push(st.border.take().unwrap_or_default());
                    }
                    "left" | "right" | "top" | "bottom" => st.edge = None,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    dict
}

fn on_style_event(
    dict: &mut StyleDict,
    st: &mut DictState,
    e: &BytesStart,
    is_empty: bool,
    theme: &[String],
) {
    let name = strip_ns(e.name().as_ref());
    match name.as_str() {
        // ----- section openers -----
        "fonts" => st.section = StyleSection::Fonts,
        "fills" => st.section = StyleSection::Fills,
        "borders" => st.section = StyleSection::Borders,
        "numFmts" => st.section = StyleSection::NumFmts,
        // ----- fonts -----
        "font" if st.section == StyleSection::Fonts => {
            st.font = Some(FontDef::default());
            if is_empty {
                dict.fonts.push(st.font.take().unwrap_or_default());
            }
        }
        _ if st.section == StyleSection::Fonts => {
            let Some(f) = st.font.as_mut() else { return };
            match name.as_str() {
                "b" => f.bold = elem_enabled(e),
                "i" => f.italic = elem_enabled(e),
                "u" => f.underline = attr_str(e, "val").as_deref() != Some("none"),
                "strike" => f.strikethrough = elem_enabled(e),
                "sz" => f.size = attr_f64(e, "val"),
                "name" => f.name = attr_str(e, "val"),
                "color" => f.color = parse_color_element(e, theme),
                _ => {}
            }
        }
        // ----- fills -----
        "fill" if st.section == StyleSection::Fills => {
            st.fill = Some(FillDef::default());
            if is_empty {
                dict.fills.push(None);
            }
        }
        "patternFill" if st.section == StyleSection::Fills => {
            let pattern = attr_str(e, "patternType").unwrap_or_default();
            if let Some(f) = st.fill.as_mut() {
                // `none` / `gray125` are Excel's "no fill" placeholders.
                f.solid = pattern == "solid";
                if pattern == "gray125" || pattern == "none" {
                    f.fg = None;
                }
            }
        }
        "fgColor" if st.section == StyleSection::Fills => {
            if let (Some(f), Some(c)) = (st.fill.as_mut(), parse_color_element(e, theme)) {
                f.fg = Some(c);
            }
        }
        // ----- borders -----
        "border" if st.section == StyleSection::Borders => {
            st.border = Some(CellBorders::default());
            if is_empty {
                dict.borders.push(CellBorders::default());
            }
        }
        "left" | "right" | "top" | "bottom" if st.section == StyleSection::Borders => {
            let edge = match name.as_str() {
                "left" => BorderEdge::Left,
                "right" => BorderEdge::Right,
                "top" => BorderEdge::Top,
                _ => BorderEdge::Bottom,
            };
            if let Some(b) = st.border.as_mut()
                && let Some(style) = attr_str(e, "style")
                && style != "none"
            {
                let border = Border {
                    style: border_style_from(&style),
                    color: "#000000".to_string(),
                };
                match edge {
                    BorderEdge::Left => b.left = Some(border),
                    BorderEdge::Right => b.right = Some(border),
                    BorderEdge::Top => b.top = Some(border),
                    BorderEdge::Bottom => b.bottom = Some(border),
                }
            }
            if !is_empty {
                st.edge = Some(edge);
            }
        }
        "color" if st.section == StyleSection::Borders => {
            if let (Some(b), Some(edge), Some(c)) = (
                st.border.as_mut(),
                st.edge,
                parse_color_element(e, theme),
            ) {
                let target = match edge {
                    BorderEdge::Left => b.left.as_mut(),
                    BorderEdge::Right => b.right.as_mut(),
                    BorderEdge::Top => b.top.as_mut(),
                    BorderEdge::Bottom => b.bottom.as_mut(),
                };
                if let Some(t) = target {
                    t.color = c;
                }
            }
        }
        // ----- custom number formats -----
        "numFmt" if st.section == StyleSection::NumFmts => {
            if let (Some(id), Some(code)) =
                (attr_u32(e, "numFmtId"), attr_str(e, "formatCode"))
            {
                dict.num_fmts.insert(id, code);
            }
        }
        _ => {}
    }
}

/// Pass 2 of `xl/styles.xml`: resolve each `<cellXfs><xf>` to a `CellFormat`.
///
/// Relies on the schema ordering (fonts/fills/borders/numFmts precede cellXfs),
/// which every mainstream writer honours.
fn parse_cell_xfs(xml: &str, dict: &StyleDict) -> Vec<CellFormat> {
    let mut out = Vec::new();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut in_cell_xfs = false;
    let mut cur: Option<XfDef> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = strip_ns(e.name().as_ref());
                match name.as_str() {
                    "cellXfs" => in_cell_xfs = true,
                    "xf" if in_cell_xfs => cur = Some(xf_from_element(&e)),
                    "alignment" if in_cell_xfs => {
                        if let Some(c) = cur.as_mut() {
                            c.alignment = Some(align_from_element(&e));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = strip_ns(e.name().as_ref());
                if in_cell_xfs {
                    match name.as_str() {
                        "xf" => {
                            let xf = xf_from_element(&e);
                            out.push(resolve_xf(&xf, dict));
                        }
                        "alignment" => {
                            if let Some(c) = cur.as_mut() {
                                c.alignment = Some(align_from_element(&e));
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = strip_ns(e.name().as_ref());
                if name == "cellXfs" {
                    in_cell_xfs = false;
                } else if in_cell_xfs && name == "xf" {
                    let xf = cur.take().unwrap_or_default();
                    out.push(resolve_xf(&xf, dict));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn xf_from_element(e: &BytesStart) -> XfDef {
    XfDef {
        num_fmt_id: attr_u32(e, "numFmtId").unwrap_or(0),
        font_id: attr_u32(e, "fontId").unwrap_or(0),
        fill_id: attr_u32(e, "fillId").unwrap_or(0),
        border_id: attr_u32(e, "borderId").unwrap_or(0),
        alignment: None,
    }
}

fn align_from_element(e: &BytesStart) -> AlignDef {
    let h = match attr_str(e, "horizontal").as_deref() {
        Some("left") => Some(HAlign::Left),
        Some("center") | Some("centerContinuous") => Some(HAlign::Center),
        Some("right") => Some(HAlign::Right),
        // `general` / `justify` / `fill` / `distributed`: keep our default.
        _ => None,
    };
    let v = match attr_str(e, "vertical").as_deref() {
        Some("top") => Some(VAlign::Top),
        Some("center") => Some(VAlign::Middle),
        Some("bottom") => Some(VAlign::Bottom),
        _ => None,
    };
    AlignDef {
        h,
        v,
        wrap: attr_truthy(e, "wrapText"),
        rotation: attr_str(e, "textRotation")
            .and_then(|s| s.trim().parse::<i16>().ok())
            .unwrap_or(0),
        indent: attr_u32(e, "indent").unwrap_or(0).min(u8::MAX as u32) as u8,
    }
}

fn resolve_xf(xf: &XfDef, dict: &StyleDict) -> CellFormat {
    let mut cf = CellFormat::default();

    if let Some(f) = dict.fonts.get(xf.font_id as usize) {
        cf.bold = f.bold;
        cf.italic = f.italic;
        cf.underline = f.underline;
        cf.strikethrough = f.strikethrough;
        if let Some(sz) = f.size {
            cf.font_size = sz;
        }
        if let Some(ref n) = f.name {
            cf.font_family = n.clone();
        }
        cf.font_color = f.color.clone();
    }
    if let Some(fill) = dict.fills.get(xf.fill_id as usize) {
        cf.bg_color = fill.clone();
    }
    if let Some(b) = dict.borders.get(xf.border_id as usize) {
        cf.borders = b.clone();
    }
    cf.number_format = dict
        .num_fmts
        .get(&xf.num_fmt_id)
        .cloned()
        .or_else(|| builtin_num_fmt(xf.num_fmt_id).map(|s| s.to_string()));

    if let Some(a) = &xf.alignment {
        if let Some(h) = a.h.clone() {
            cf.h_align = h;
        }
        if let Some(v) = a.v.clone() {
            cf.v_align = v;
        }
        if a.wrap {
            cf.text_wrap = TextWrap::Wrap;
        }
        cf.text_rotation = a.rotation;
        cf.indent = a.indent;
    }

    cf
}

/// Does a resolved format carry any *visible* styling worth attaching?
///
/// Thin wrapper over [`CellFormat::is_notable`] so the reader and the writer
/// cannot drift apart on "what counts as styled".
fn format_is_notable(cf: &CellFormat) -> bool {
    cf.is_notable()
}

/// How far past the last **valued** cell a styled-but-empty cell may sit and
/// still be materialised.
///
/// A formatted empty cell is how Excel represents a formatted-but-empty area
/// (a bordered empty row, the empty right half of a bordered table, …), so it
/// has to survive the import. But a stray styled cell far away would inflate
/// `used_range` and bloat memory, so cells beyond the content plus a small
/// margin are treated as "outside the table" and dropped.
/// 8 matches `apply_format`'s own clamp (`used_range + 8`) in the app.
const STYLE_MARGIN: u32 = 8;

/// Bottom-right corner of the cells that actually **hold a value** (0-based).
///
/// Note this is *not* `Sheet::used_range()`: the latter is the max key in the
/// cell map, which after this import includes the very styled-but-empty cells
/// we are about to create.
fn valued_bounds(sheet: &Sheet) -> (u32, u32) {
    let mut max_row = 0;
    let mut max_col = 0;
    for (&(row, col), cell) in sheet.cells() {
        if matches!(cell.value, CellValue::Empty) {
            continue;
        }
        max_row = max_row.max(row);
        max_col = max_col.max(col);
    }
    (max_row, max_col)
}

/// Sheet-level bits read from a worksheet XML.
#[derive(Default)]
struct SheetLayout {
    /// `(row, col)` -> `cellXfs` index (`s="N"`).
    cell_styles: HashMap<(u32, u32), usize>,
    merges: Vec<(u32, u32, u32, u32)>,
    col_widths: HashMap<u32, f64>,
    row_heights: HashMap<u32, f64>,
    hidden_rows: HashSet<u32>,
    hidden_cols: HashSet<u32>,
    tab_color: Option<String>,
}

/// Parse `"A1:C3"` (single-cell `"A1"` also accepted) to 0-based corners.
fn parse_range_ref(range: &str) -> Option<(u32, u32, u32, u32)> {
    let (start, end) = match range.split_once(':') {
        Some((a, b)) => (a, b),
        None => (range, range),
    };
    let (sr, sc) = parse_a1_ref(start.trim())?;
    let (er, ec) = parse_a1_ref(end.trim())?;
    if er < sr || ec < sc {
        return None;
    }
    Some((sr, sc, er, ec))
}

/// Parse the layout parts of a worksheet XML: cell style indices, merged
/// regions, column widths, row heights, hidden rows/columns and tab colour.
fn parse_sheet_layout_xml(xml: &str, theme: &[String]) -> SheetLayout {
    let mut layout = SheetLayout::default();
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut cur_row: u32 = 0;
    let mut next_col: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                on_layout_event(&mut layout, &mut cur_row, &mut next_col, &e, theme)
            }
            Ok(Event::Empty(e)) => {
                on_layout_event(&mut layout, &mut cur_row, &mut next_col, &e, theme)
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    layout
}

fn on_layout_event(
    layout: &mut SheetLayout,
    cur_row: &mut u32,
    next_col: &mut u32,
    e: &BytesStart,
    theme: &[String],
) {
    let name = strip_ns(e.name().as_ref());
    match name.as_str() {
        "row" => {
            *cur_row = attr_u32(e, "r").unwrap_or(1).saturating_sub(1);
            *next_col = 0;
            let custom = match attr_str(e, "customHeight") {
                Some(_) => attr_truthy(e, "customHeight"),
                None => true,
            };
            if let Some(h) = attr_f64(e, "ht")
                && custom
                && h > 0.0
            {
                layout.row_heights.insert(*cur_row, h);
            }
            if attr_truthy(e, "hidden") {
                layout.hidden_rows.insert(*cur_row);
            }
        }
        "c" => {
            let (row, col) = match attr_str(e, "r").and_then(|s| parse_a1_ref(&s)) {
                Some(rc) => rc,
                // Legal to omit `r`: then the column is positional.
                None => (*cur_row, *next_col),
            };
            *next_col = col.saturating_add(1);
            if let Some(s) = attr_u32(e, "s") {
                layout.cell_styles.insert((row, col), s as usize);
            }
        }
        "col" => {
            let min = attr_u32(e, "min").unwrap_or(0);
            let max = attr_u32(e, "max").unwrap_or(0);
            if min == 0 || max < min {
                return;
            }
            let span = (max - min + 1).min(MAX_COL_SPAN);
            let width = attr_f64(e, "width");
            let custom = match attr_str(e, "customWidth") {
                Some(_) => attr_truthy(e, "customWidth"),
                None => true,
            };
            let hidden = attr_truthy(e, "hidden");
            for i in 0..span {
                let col = min - 1 + i;
                if let Some(w) = width
                    && custom
                    && w > 0.0
                {
                    layout.col_widths.insert(col, w);
                }
                if hidden {
                    layout.hidden_cols.insert(col);
                }
            }
        }
        "mergeCell" => {
            if let Some((sr, sc, er, ec)) = attr_str(e, "ref").and_then(|r| parse_range_ref(&r))
                && (sr != er || sc != ec)
            {
                layout.merges.push((sr, sc, er, ec));
            }
        }
        "tabColor" => {
            if let Some(c) = parse_color_element(e, theme) {
                layout.tab_color = Some(c);
            }
        }
        _ => {}
    }
}

/// Read cell styles and sheet layout from the raw xlsx bytes and apply them.
///
/// Styles are attached only to cells that already exist (value or formula):
/// materialising the styled-but-empty cells Excel loves to emit would inflate
/// `used_range` and disturb every range-based table operation.
fn apply_styles_and_layout_from_bytes(bytes: &[u8], workbook: &mut Workbook) -> Result<()> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|e| IoError::XlsxRead(format!("zip error: {}", e)))?;

    let theme = match read_zip_entry_string(&mut archive, "xl/theme/theme1.xml") {
        Ok(xml) => parse_theme_colors(&xml),
        Err(_) => Vec::new(),
    };
    let styles = match read_zip_entry_string(&mut archive, "xl/styles.xml") {
        Ok(xml) => {
            let dict = parse_style_dict(&xml, &theme);
            parse_cell_xfs(&xml, &dict)
        }
        Err(_) => Vec::new(),
    };

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml")?;
    let sheet_to_rid = parse_sheet_rid_map(&workbook_xml);
    let rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rid_to_target = parse_rid_target_map(&rels_xml);

    for sheet_name in workbook.sheet_names() {
        let Some(rid) = sheet_to_rid.get(sheet_name.as_str()) else {
            continue;
        };
        let Some(target) = rid_to_target.get(rid.as_str()) else {
            continue;
        };
        let xml_path = format!("xl/{}", target);
        let Ok(sheet_xml) = read_zip_entry_string(&mut archive, &xml_path) else {
            continue;
        };
        let layout = parse_sheet_layout_xml(&sheet_xml, &theme);

        let Ok(sheet) = workbook.get_sheet_mut(sheet_name.as_str()) else {
            continue;
        };

        // Bottom-right corner of the *valued* cells: the yardstick for which
        // styled-but-empty cells are near enough to the table to keep.
        let (valued_max_row, valued_max_col) = valued_bounds(sheet);
        for (&(row, col), &idx) in &layout.cell_styles {
            let Some(cf) = styles.get(idx) else { continue };
            if !format_is_notable(cf) {
                continue;
            }
            if let Some(cell) = sheet.get_cell_mut(row, col) {
                cell.format = cf.clone();
            } else if row <= valued_max_row.saturating_add(STYLE_MARGIN)
                && col <= valued_max_col.saturating_add(STYLE_MARGIN)
            {
                // 有样式、没值：必须物化成 `Empty` 占位。不物化的话，边框/底色这些
                // “靠格子存在才画得出来”的东西会整片丢失——Excel 正是用这种空格子
                // 表示“格式化过但没内容”的区域（表格框线的下半截、留白、模板）。
                sheet.set_cell(
                    row,
                    col,
                    Cell {
                        value: CellValue::Empty,
                        format: cf.clone(),
                        ..Default::default()
                    },
                );
            }
        }

        for &(sr, sc, er, ec) in &layout.merges {
            sheet.add_merged_region_from_import(sr, sc, er, ec);
        }
        sheet.col_widths.extend(layout.col_widths);
        sheet.row_heights.extend(layout.row_heights);
        sheet.hidden_rows.extend(layout.hidden_rows);
        sheet.hidden_cols.extend(layout.hidden_cols);
        if let Some(tc) = layout.tab_color {
            sheet.tab_color = Some(tc);
        }
    }

    Ok(())
}

/// Parse an A1-style cell reference (e.g. "A1", "AB123") to 0-based (row, col).
///
/// Returns `None` if the reference is malformed.
fn parse_a1_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let first_digit = cell_ref.find(|c: char| c.is_ascii_digit())?;
    if first_digit == 0 {
        return None;
    }
    let col_part = &cell_ref[..first_digit];
    let row_part = &cell_ref[first_digit..];

    // Column letters to 0-based index.
    let mut col: u32 = 0;
    for ch in col_part.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
    }
    col = col.checked_sub(1)?; // 1-based -> 0-based

    let row: u32 = row_part.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, col)) // 1-based -> 0-based
}

/// Strip namespace prefix from an XML name (e.g. `"r:id"` -> `"id"`).
fn strip_ns(full: &[u8]) -> String {
    let s = String::from_utf8_lossy(full);
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

/// Convert calamine `Data` enum to our `CellValue`.
fn calamine_data_to_cell_value(data: &Data) -> CellValue {
    match data {
        Data::Empty => CellValue::Empty,
        Data::String(s) => CellValue::Text(s.clone()),
        Data::Float(f) => CellValue::Number(*f),
        Data::Int(i) => CellValue::Number(*i as f64),
        Data::Bool(b) => CellValue::Boolean(*b),
        Data::Error(e) => {
            let cell_error = match e {
                calamine::CellErrorType::Div0 => CellError::DivZero,
                calamine::CellErrorType::NA => CellError::NA,
                calamine::CellErrorType::Name => CellError::Name,
                calamine::CellErrorType::Null => CellError::Null,
                calamine::CellErrorType::Num => CellError::Num,
                calamine::CellErrorType::Ref => CellError::Ref,
                calamine::CellErrorType::Value => CellError::Value,
                calamine::CellErrorType::GettingData => CellError::NA,
            };
            CellValue::Error(cell_error)
        }
        Data::DateTime(dt) => {
            // ExcelDateTime stores a serial number. Convert to ISO 8601 string.
            let serial = dt.as_f64();
            CellValue::Date(excel_serial_to_iso(serial))
        }
        Data::DateTimeIso(s) => CellValue::Date(s.clone()),
        Data::DurationIso(s) => CellValue::Text(s.clone()),
    }
}

/// Convert an Excel serial date number to an ISO 8601 date string.
///
/// Excel uses a serial date system where 1 = 1900-01-01.
/// Due to the Lotus 1-2-3 bug, Excel incorrectly treats 1900 as a leap year,
/// so dates >= 60 are off by one day.
fn excel_serial_to_iso(serial: f64) -> String {
    // Number of days from Excel epoch (1899-12-30) to Unix epoch (1970-01-01)
    const EXCEL_EPOCH_OFFSET: i64 = 25569;
    const SECONDS_PER_DAY: i64 = 86400;

    let days = serial as i64;
    let fraction = serial - days as f64;

    // Convert to Unix timestamp
    let unix_days = days - EXCEL_EPOCH_OFFSET;
    let total_seconds = unix_days * SECONDS_PER_DAY + (fraction * SECONDS_PER_DAY as f64) as i64;

    // Simple date calculation from Unix timestamp
    let (year, month, day, hour, minute, second) = unix_timestamp_to_date(total_seconds);

    if hour == 0 && minute == 0 && second == 0 {
        format!("{:04}-{:02}-{:02}", year, month, day)
    } else {
        format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            year, month, day, hour, minute, second
        )
    }
}

/// Convert an ISO date string back to an Excel serial date number.
///
/// Supports `"YYYY-MM-DD"` and `"YYYY-MM-DDThh:mm:ss"` formats.
pub(crate) fn iso_to_excel_serial(iso: &str) -> Option<f64> {
    // Parse YYYY-MM-DD
    let parts: Vec<&str> = iso.split('T').collect();
    let date_part = parts.first()?;
    let date_fields: Vec<&str> = date_part.split('-').collect();
    if date_fields.len() != 3 {
        return None;
    }
    let year: i32 = date_fields[0].parse().ok()?;
    let month: u32 = date_fields[1].parse().ok()?;
    let day: u32 = date_fields[2].parse().ok()?;

    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }

    // Parse optional time part.
    let (hour, minute, second) = if parts.len() > 1 {
        let time_fields: Vec<&str> = parts[1].split(':').collect();
        let h: u32 = time_fields
            .first()
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);
        let m: u32 = time_fields.get(1).and_then(|s| s.parse().ok()).unwrap_or(0);
        let s: u32 = time_fields.get(2).and_then(|s| s.parse().ok()).unwrap_or(0);
        (h, m, s)
    } else {
        (0, 0, 0)
    };

    // Convert to Unix timestamp, then to Excel serial.
    let unix_ts = date_to_unix_timestamp(year, month, day, hour, minute, second);
    const EXCEL_EPOCH_OFFSET: i64 = 25569;
    const SECONDS_PER_DAY: i64 = 86400;

    let serial = (unix_ts as f64) / (SECONDS_PER_DAY as f64) + EXCEL_EPOCH_OFFSET as f64;
    Some(serial)
}

/// Convert date components to a Unix timestamp.
fn date_to_unix_timestamp(year: i32, month: u32, day: u32, hour: u32, min: u32, sec: u32) -> i64 {
    let seconds_per_day: i64 = 86400;

    // Days from 1970-01-01 to the start of the given year.
    let mut total_days: i64 = 0;
    if year >= 1970 {
        for y in 1970..year {
            total_days += if is_leap_year(y) { 366 } else { 365 };
        }
    } else {
        for y in year..1970 {
            total_days -= if is_leap_year(y) { 366 } else { 365 };
        }
    }

    // Days from start of year to start of month.
    let leap = is_leap_year(year);
    let month_days: [u32; 12] = if leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };
    for &md in month_days.iter().take((month - 1) as usize) {
        total_days += md as i64;
    }
    total_days += (day - 1) as i64;

    total_days * seconds_per_day + (hour as i64) * 3600 + (min as i64) * 60 + sec as i64
}

/// Convert a Unix timestamp to (year, month, day, hour, minute, second).
fn unix_timestamp_to_date(timestamp: i64) -> (i32, u32, u32, u32, u32, u32) {
    let seconds_in_day = 86400i64;
    let mut days = timestamp / seconds_in_day;
    let mut remaining_seconds = (timestamp % seconds_in_day) as u32;
    if timestamp < 0 && remaining_seconds > 0 {
        days -= 1;
        remaining_seconds = (seconds_in_day + (timestamp % seconds_in_day)) as u32;
    }

    let hour = remaining_seconds / 3600;
    remaining_seconds %= 3600;
    let minute = remaining_seconds / 60;
    let second = remaining_seconds % 60;

    // Days since 1970-01-01
    let mut year = 1970i32;
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if days < days_in_year {
            break;
        }
        days -= days_in_year;
        year += 1;
    }

    let leap = is_leap_year(year);
    let month_days: [i64; 12] = if leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 0u32;
    for (i, &md) in month_days.iter().enumerate() {
        if days < md {
            month = i as u32 + 1;
            break;
        }
        days -= md;
    }
    if month == 0 {
        month = 12;
    }

    let day = days as u32 + 1;
    (year, month, day, hour, minute, second)
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_calamine_empty() {
        assert_eq!(calamine_data_to_cell_value(&Data::Empty), CellValue::Empty);
    }

    #[test]
    fn test_calamine_string() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::String("hello".into())),
            CellValue::Text("hello".into())
        );
    }

    #[test]
    fn test_calamine_float() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::Float(42.5)),
            CellValue::Number(42.5)
        );
    }

    #[test]
    fn test_calamine_bool() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::Bool(true)),
            CellValue::Boolean(true)
        );
    }

    #[test]
    fn test_excel_serial_date() {
        // 44197 = 2021-01-01 in Excel serial dates
        let iso = excel_serial_to_iso(44197.0);
        assert_eq!(iso, "2021-01-01");
    }

    #[test]
    fn test_excel_serial_datetime() {
        // 44197.5 = 2021-01-01 12:00:00
        let iso = excel_serial_to_iso(44197.5);
        assert_eq!(iso, "2021-01-01T12:00:00");
    }

    #[test]
    fn test_iso_to_excel_serial_date() {
        let serial = iso_to_excel_serial("2021-01-01").unwrap();
        // Should round-trip to the same value.
        assert!((serial - 44197.0).abs() < 0.001);
    }

    #[test]
    fn test_iso_to_excel_serial_datetime() {
        let serial = iso_to_excel_serial("2021-01-01T12:00:00").unwrap();
        assert!((serial - 44197.5).abs() < 0.001);
    }

    #[test]
    fn test_iso_to_excel_serial_invalid() {
        assert!(iso_to_excel_serial("not-a-date").is_none());
        assert!(iso_to_excel_serial("").is_none());
    }

    #[test]
    fn test_calamine_int() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::Int(7)),
            CellValue::Number(7.0)
        );
    }

    #[test]
    fn test_calamine_error_types() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::Error(calamine::CellErrorType::Div0)),
            CellValue::Error(CellError::DivZero)
        );
        assert_eq!(
            calamine_data_to_cell_value(&Data::Error(calamine::CellErrorType::Ref)),
            CellValue::Error(CellError::Ref)
        );
        assert_eq!(
            calamine_data_to_cell_value(&Data::Error(calamine::CellErrorType::Value)),
            CellValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_calamine_duration_iso() {
        assert_eq!(
            calamine_data_to_cell_value(&Data::DurationIso("PT1H30M".into())),
            CellValue::Text("PT1H30M".into())
        );
    }

    #[test]
    fn test_parse_a1_ref_simple() {
        assert_eq!(parse_a1_ref("A1"), Some((0, 0)));
        assert_eq!(parse_a1_ref("B2"), Some((1, 1)));
        assert_eq!(parse_a1_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_a1_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_a1_ref("D5"), Some((4, 3)));
    }

    #[test]
    fn test_parse_a1_ref_invalid() {
        assert_eq!(parse_a1_ref(""), None);
        assert_eq!(parse_a1_ref("123"), None);
        assert_eq!(parse_a1_ref("A0"), None);
    }

    #[test]
    fn test_parse_sheet_rid_map() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Dashboard" sheetId="1" r:id="rId3"/>
    <sheet name="Portfolio" sheetId="2" r:id="rId4"/>
  </sheets>
</workbook>"#;
        let map = parse_sheet_rid_map(xml);
        assert_eq!(map.get("Dashboard"), Some(&"rId3".to_string()));
        assert_eq!(map.get("Portfolio"), Some(&"rId4".to_string()));
        assert_eq!(map.len(), 2);
    }

    #[test]
    fn test_parse_rid_target_map() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#;
        let map = parse_rid_target_map(xml);
        assert_eq!(map.get("rId3"), Some(&"worksheets/sheet1.xml".to_string()));
        assert_eq!(map.get("rId4"), Some(&"worksheets/sheet2.xml".to_string()));
    }

    #[test]
    fn test_parse_formulas_from_sheet_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="n"><v>100</v></c>
    </row>
    <row r="5">
      <c r="D5" s="5" t="n"><f aca="false">C5-B5</f><v>98270</v></c>
      <c r="E5" s="6" t="n"><f aca="false">D5/B5*100</f><v>10.28</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let formulas = parse_formulas_from_sheet_xml(xml);
        assert_eq!(formulas.len(), 2);
        assert_eq!(formulas[0], ("D5".to_string(), "C5-B5".to_string()));
        assert_eq!(formulas[1], ("E5".to_string(), "D5/B5*100".to_string()));
    }

    #[test]
    fn test_parse_formulas_no_formulas() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let formulas = parse_formulas_from_sheet_xml(xml);
        assert!(formulas.is_empty());
    }

    #[test]
    fn test_parse_formulas_sum_function() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="5">
      <c r="N5" t="n"><f>SUM(B5:M5)</f><v>120000</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let formulas = parse_formulas_from_sheet_xml(xml);
        assert_eq!(formulas.len(), 1);
        assert_eq!(formulas[0], ("N5".to_string(), "SUM(B5:M5)".to_string()));
    }

    /// Integration test: read the real Finance Tracker xlsx and verify formulas
    /// are extracted. This test is ignored in CI (file must be present locally).
    #[test]
    fn test_read_xlsx_with_formulas_real_file() {
        let home = std::env::var("HOME").unwrap_or_default();
        let path =
            std::path::PathBuf::from(format!("{}/Downloads/Finance_Tracker_FY2025-26.xlsx", home));
        if !path.exists() {
            eprintln!("skipping: test file not found at {:?}", path);
            return;
        }
        let wb = read_xlsx(&path).expect("should read xlsx");

        // Dashboard sheet should exist.
        let dashboard = wb
            .get_sheet("Dashboard")
            .expect("should have Dashboard sheet");

        // D5 should have formula "C5-B5" (0-based: row 4, col 3).
        let cell_d5 = dashboard.get_cell(4, 3).expect("D5 should exist");
        assert!(
            cell_d5.formula.is_some(),
            "D5 should have a formula, got {:?}",
            cell_d5
        );
        assert_eq!(cell_d5.formula.as_deref(), Some("C5-B5"));

        // E5 should have formula "D5/B5*100" (0-based: row 4, col 4).
        let cell_e5 = dashboard.get_cell(4, 4).expect("E5 should exist");
        assert_eq!(cell_e5.formula.as_deref(), Some("D5/B5*100"));

        // Income sheet — N5 should have SUM formula.
        let income = wb.get_sheet("Income").expect("should have Income sheet");
        let cell_n5 = income.get_cell(4, 13).expect("N5 should exist");
        assert!(cell_n5.formula.is_some(), "Income!N5 should have a formula");
        assert_eq!(cell_n5.formula.as_deref(), Some("SUM(B5:M5)"));
    }
}

/// Tests for the styles/layout import pass added on top of the value pass.
#[cfg(test)]
mod style_import_tests {
    use super::*;
    use std::io::Write as _;

    // -------------------------------------------------------------------
    // Fixtures: hand-written xlsx parts, independent of rust_xlsxwriter.
    // Using foreign XML on purpose — the point is reading files we did not
    // write ourselves.
    // -------------------------------------------------------------------

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#;

    const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

    const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

    /// Deliberately non-standard palette: slot order in `<a:clrScheme>` is
    /// dk1/lt1/dk2/lt2/…, but `theme="N"` indexes lt1/dk1/lt2/dk2/….
    const THEME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="A00001"/></a:accent1><a:accent2><a:srgbClr val="A00002"/></a:accent2><a:accent3><a:srgbClr val="A00003"/></a:accent3><a:accent4><a:srgbClr val="A00004"/></a:accent4><a:accent5><a:srgbClr val="A00005"/></a:accent5><a:accent6><a:srgbClr val="A00006"/></a:accent6><a:hlink><a:srgbClr val="A00007"/></a:hlink><a:folHlink><a:srgbClr val="A00008"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#;

    /// 5 `cellXfs` entries: default / rich header / money / custom 0.000 / percent.
    const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Calibri"/><color theme="1"/></font><font><sz val="11"/><name val="Calibri"/><color rgb="FFFF0000"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FF0000FF"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="4" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

    /// Layout probe: does not go through calamine, only through the layout parser.
    const LAYOUT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF00B050"/></sheetPr><cols><col min="2" max="2" width="24.5" customWidth="1"/><col min="4" max="5" width="9" customWidth="1" hidden="1"/></cols><sheetData><row r="1" ht="30" customHeight="1"><c r="A1" s="1" t="inlineStr"><is><t>标题</t></is></c></row><row r="2" hidden="1"><c r="A2" s="3" t="n"><v>1234.5</v></c></row><row r="3"><c s="1" t="n"><v>7</v></c><c r="C4" s="4" t="n"><v>8</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells></worksheet>"#;

    /// End-to-end sheet. Every `<c>` carries an explicit `r` so the test does
    /// not depend on how calamine treats positional cells.
    const E2E_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF00B050"/></sheetPr><dimension ref="A1:D4"/><cols><col min="2" max="2" width="24.5" customWidth="1"/><col min="4" max="5" width="9" customWidth="1" hidden="1"/></cols><sheetData><row r="1" ht="30" customHeight="1"><c r="A1" s="1" t="inlineStr"><is><t>标题</t></is></c><c r="B1" s="2" t="n"><v>10</v></c><c r="D1" s="1"/></row><row r="2" hidden="1"><c r="A2" s="3" t="n"><v>1234.5</v></c><c r="B2" s="0" t="n"><v>5</v></c></row><row r="3"><c r="A3" s="4" t="n"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells></worksheet>"#;

    fn build_xlsx(parts: &[(&str, &str)]) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, body) in parts {
                zw.start_file(*name, opts).unwrap();
                zw.write_all(body.as_bytes()).unwrap();
            }
            zw.finish().unwrap();
        }
        buf
    }

    fn foreign_xlsx() -> Vec<u8> {
        build_xlsx(&[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
            ("xl/styles.xml", STYLES_XML),
            ("xl/theme/theme1.xml", THEME_XML),
            ("xl/worksheets/sheet1.xml", E2E_SHEET_XML),
        ])
    }

    // -------------------------------------------------------------------
    // Small helpers
    // -------------------------------------------------------------------

    #[test]
    fn normalize_hex_strips_alpha_and_uppercases() {
        assert_eq!(normalize_hex("FFFFC000"), "#FFC000");
        assert_eq!(normalize_hex("ffc000"), "#FFC000");
        assert_eq!(normalize_hex("FF0000"), "#FF0000");
        assert_eq!(normalize_hex("#00B050"), "#00B050");
    }

    #[test]
    fn builtin_num_fmt_table() {
        assert_eq!(builtin_num_fmt(0), None, "General 不算格式");
        assert_eq!(builtin_num_fmt(4), Some("#,##0.00"));
        assert_eq!(builtin_num_fmt(9), Some("0%"));
        assert_eq!(builtin_num_fmt(14), Some("m/d/yyyy"));
        assert_eq!(builtin_num_fmt(164), None, "自定义格式要去 numFmts 查");
    }

    #[test]
    fn border_style_names_map_to_variants() {
        assert_eq!(border_style_from("thin"), BorderStyle::Thin);
        assert_eq!(border_style_from("medium"), BorderStyle::Medium);
        assert_eq!(border_style_from("thick"), BorderStyle::Thick);
        assert_eq!(border_style_from("double"), BorderStyle::Double);
        assert_eq!(border_style_from("hair"), BorderStyle::Thin, "hairline 是实线，不是虚线");
        assert_eq!(border_style_from("dotted"), BorderStyle::Dotted);
        assert_eq!(border_style_from("dashDot"), BorderStyle::Dashed);
        assert_eq!(border_style_from("none"), BorderStyle::None);
    }

    #[test]
    fn parse_range_ref_accepts_single_cell_and_rejects_inverted() {
        assert_eq!(parse_range_ref("A1:C3"), Some((0, 0, 2, 2)));
        assert_eq!(parse_range_ref("B2"), Some((1, 1, 1, 1)));
        assert_eq!(parse_range_ref("C3:A1"), None);
        assert_eq!(parse_range_ref("nonsense"), None);
    }

    // -------------------------------------------------------------------
    // Theme colours
    // -------------------------------------------------------------------

    #[test]
    fn theme_slots_are_remapped_to_theme_index_order() {
        let theme = parse_theme_colors(THEME_XML);
        assert_eq!(theme.len(), 12);
        assert_eq!(theme[0], "#FEFEFE", "index 0 是 lt1，不是文档里第一个的 dk1");
        assert_eq!(theme[1], "#111111", "index 1 是 dk1（走 sysClr lastClr）");
        assert_eq!(theme[2], "#EEEEEE");
        assert_eq!(theme[3], "#222222");
        assert_eq!(theme[4], "#A00001");
        assert_eq!(theme[10], "#A00007", "hlink");
    }

    #[test]
    fn missing_theme_falls_back_to_office_defaults() {
        let theme = parse_theme_colors("<a:theme/>");
        assert_eq!(theme[0], "#FFFFFF");
        assert_eq!(theme[1], "#000000");
        assert_eq!(theme[4], "#4472C4");
    }

    // -------------------------------------------------------------------
    // styles.xml
    // -------------------------------------------------------------------

    #[test]
    fn style_dict_reads_fonts_fills_borders_and_custom_num_fmts() {
        let theme = parse_theme_colors(THEME_XML);
        let dict = parse_style_dict(STYLES_XML, &theme);
        assert_eq!(dict.fonts.len(), 3);
        assert_eq!(dict.fills.len(), 3);
        assert_eq!(dict.borders.len(), 2);
        assert_eq!(dict.num_fmts.get(&164).map(String::as_str), Some("0.000"));

        assert!(dict.fonts[1].bold);
        assert_eq!(dict.fonts[1].size, Some(12.0));
        assert_eq!(dict.fonts[1].color.as_deref(), Some("#111111"));
        assert_eq!(dict.fonts[2].color.as_deref(), Some("#FF0000"));

        assert_eq!(dict.fills[0], None, "patternType=none → 无填充");
        assert_eq!(dict.fills[1], None, "patternType=gray125 → 无填充");
        assert_eq!(dict.fills[2].as_deref(), Some("#FFC000"));

        let bottom = dict.borders[1].bottom.as_ref().expect("thin bottom");
        assert_eq!(bottom.style, BorderStyle::Thin);
        assert_eq!(bottom.color, "#0000FF");
        assert!(dict.borders[0].bottom.is_none(), "无 style 的边不产生边框");
    }

    #[test]
    fn cell_xfs_resolve_to_notable_formats() {
        let theme = parse_theme_colors(THEME_XML);
        let dict = parse_style_dict(STYLES_XML, &theme);
        let xfs = parse_cell_xfs(STYLES_XML, &dict);
        assert_eq!(xfs.len(), 5, "cellStyleXfs 不应混进来");

        assert!(
            !format_is_notable(&xfs[0]),
            "默认 xf（Calibri 11、无填充/边框）不应算有样式"
        );

        let rich = &xfs[1];
        assert!(rich.bold);
        assert_eq!(rich.font_size, 12.0);
        assert_eq!(rich.font_family, "Calibri");
        assert_eq!(rich.font_color.as_deref(), Some("#111111"));
        assert_eq!(rich.bg_color.as_deref(), Some("#FFC000"));
        assert_eq!(rich.h_align, HAlign::Center);
        assert_eq!(rich.v_align, VAlign::Middle);
        assert_eq!(rich.text_wrap, TextWrap::Wrap);
        assert_eq!(
            rich.borders.bottom.as_ref().map(|b| &b.style),
            Some(&BorderStyle::Thin)
        );
        assert!(format_is_notable(rich));

        assert_eq!(xfs[2].number_format.as_deref(), Some("#,##0.00"));
        assert_eq!(xfs[2].font_color.as_deref(), Some("#FF0000"));
        assert_eq!(xfs[3].number_format.as_deref(), Some("0.000"));
        assert_eq!(xfs[4].number_format.as_deref(), Some("0%"));
    }

    // -------------------------------------------------------------------
    // worksheet layout
    // -------------------------------------------------------------------

    #[test]
    fn layout_parser_reads_cols_rows_merges_hidden_and_tab_color() {
        let theme = parse_theme_colors(THEME_XML);
        let layout = parse_sheet_layout_xml(LAYOUT_XML, &theme);

        assert_eq!(layout.col_widths.get(&1), Some(&24.5));
        assert_eq!(layout.col_widths.get(&3), Some(&9.0));
        assert_eq!(layout.col_widths.get(&4), Some(&9.0));
        assert!(layout.hidden_cols.contains(&3) && layout.hidden_cols.contains(&4));

        assert_eq!(layout.row_heights.get(&0), Some(&30.0));
        assert!(layout.hidden_rows.contains(&1));
        assert!(!layout.row_heights.contains_key(&2), "无 customHeight 不产生高度");

        assert_eq!(layout.merges, vec![(0, 0, 0, 2)]);
        assert_eq!(layout.tab_color.as_deref(), Some("#00B050"));

        assert_eq!(layout.cell_styles.get(&(0, 0)), Some(&1));
        assert_eq!(layout.cell_styles.get(&(1, 0)), Some(&3));
        assert_eq!(
            layout.cell_styles.get(&(2, 0)),
            Some(&1),
            "省略 r 的 <c> 按位置补列"
        );
        assert_eq!(layout.cell_styles.get(&(3, 2)), Some(&4));
    }

    #[test]
    fn col_span_is_capped() {
        let xml = r#"<worksheet><cols><col min="1" max="16384" width="8" customWidth="1"/></cols></worksheet>"#;
        let layout = parse_sheet_layout_xml(xml, &[]);
        assert_eq!(layout.col_widths.len(), MAX_COL_SPAN as usize);
    }

    #[test]
    fn degenerate_and_duplicate_merges_are_skipped() {
        let xml = r#"<worksheet><mergeCells count="3"><mergeCell ref="A1"/><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells></worksheet>"#;
        let layout = parse_sheet_layout_xml(xml, &[]);
        assert_eq!(layout.merges, vec![(0, 0, 1, 1), (1, 1, 2, 2)], "单格不算合并");
    }

    // -------------------------------------------------------------------
    // end to end
    // -------------------------------------------------------------------

    #[test]
    fn reads_foreign_xlsx_styles_layout_and_merges() {
        let wb = read_xlsx_from_bytes(&foreign_xlsx()).expect("should read");
        let sheet = wb.get_sheet("Sheet1").expect("Sheet1 应存在");

        let a1 = sheet.get_cell(0, 0).expect("A1");
        assert_eq!(a1.value, CellValue::Text("标题".into()));
        assert!(a1.format.bold);
        assert_eq!(a1.format.font_size, 12.0);
        assert_eq!(a1.format.font_family, "Calibri");
        assert_eq!(a1.format.font_color.as_deref(), Some("#111111"));
        assert_eq!(a1.format.bg_color.as_deref(), Some("#FFC000"));
        assert_eq!(a1.format.h_align, HAlign::Center);
        assert_eq!(a1.format.v_align, VAlign::Middle);
        assert_eq!(a1.format.text_wrap, TextWrap::Wrap);
        let bottom = a1.format.borders.bottom.as_ref().expect("下边框");
        assert_eq!(bottom.style, BorderStyle::Thin);
        assert_eq!(bottom.color, "#0000FF");

        let b1 = sheet.get_cell(0, 1).expect("B1");
        assert_eq!(b1.format.number_format.as_deref(), Some("#,##0.00"));
        assert_eq!(b1.format.font_color.as_deref(), Some("#FF0000"));

        let a2 = sheet.get_cell(1, 0).expect("A2");
        assert_eq!(a2.format.number_format.as_deref(), Some("0.000"));

        let a3 = sheet.get_cell(2, 0).expect("A3");
        assert_eq!(a3.format.number_format.as_deref(), Some("0%"));

        assert_eq!(
            sheet.get_cell(1, 1).expect("B2").format,
            CellFormat::default(),
            "普通单元格保持默认格式，不因 Calibri 而加壳"
        );
        // D1 在文件里是 `<c r="D1" s="1"/>`：有样式、没值。**必须物化**——
        // Excel 就是用这种格子表示“格式化过但没内容”的区域，丢了的话表格框线、
        // 底色会整片缺失（用户报「边框渲染的不对」）。
        let d1 = sheet.get_cell(0, 3).expect("D1 应被物化成空值占位");
        assert_eq!(d1.value, CellValue::Empty, "占位必须是空值，不能在表里凭空多出内容");
        assert!(d1.format.bold && d1.format.bg_color.is_some(), "样式要跟着一起留下");
        assert!(d1.format.borders.bottom.is_some(), "边框是用户看得见的那部分");

        assert!(
            sheet
                .merged_regions()
                .iter()
                .any(|r| r.start_row == 0 && r.start_col == 0 && r.end_row == 0 && r.end_col == 2),
            "A1:C1 合并应被导入"
        );
        assert_eq!(sheet.col_widths.get(&1), Some(&24.5));
        assert_eq!(sheet.row_heights.get(&0), Some(&30.0));
        assert!(sheet.hidden_rows.contains(&1));
        assert!(sheet.hidden_cols.contains(&3) && sheet.hidden_cols.contains(&4));
        assert_eq!(sheet.tab_color.as_deref(), Some("#00B050"));
    }

    #[test]
    fn styled_empty_cells_far_outside_the_table_are_not_materialised() {
        // 表格之外的“幻影样式格”（Excel 里常见于曾整行整列刷过格式的文件）不物化——
        // 否则一个远离数据的 A999 能把 used_range 拉到第 999 行。
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c><c r="Z1" s="1"/></row><row r="900"><c r="A900" s="1"/></row></sheetData></worksheet>"#;
        let bytes = build_xlsx(&[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/styles.xml", STYLES_XML),
        ]);
        let wb = read_xlsx_from_bytes(&bytes).expect("should read");
        let sheet = wb.get_sheet("Sheet1").unwrap();
        assert!(sheet.get_cell(899, 0).is_none(), "远处的样式格不物化（A900）");
        assert!(sheet.get_cell(0, 25).is_none(), "远处的样式格不物化（Z1）");
        // 靠得近的还得留下。
        assert!(sheet.get_cell(0, 0).is_some());
    }

    /// **不变量**：导入 → 导出 → 再导入，模型不得变样。
    ///
    /// 这是这一片区的“系统性闸门”：逐个 bug 写断言只能卡住已知的坑，
    /// 而“读得进、写不回”这类不对称丢样式的问题（框线丢一半就是这样），
    /// 只要两端对不上就会在这里暴露。
    #[test]
    fn import_export_import_is_stable() {
        let first = read_xlsx_from_bytes(&foreign_xlsx()).expect("first read");
        let bytes = crate::xlsx_writer::write_xlsx_to_buffer(&first).expect("export");
        let second = read_xlsx_from_bytes(&bytes).expect("second read");

        assert_eq!(snapshot(&second), snapshot(&first), "导入→导出→导入 不得改变模型");
    }

    /// 可比快照：格子（值 + 除去字体名/字号之外的格式）+ 合并区 + 活动表。
    ///
    /// 字体名/字号故意排除：我们写文件时用的是自己的默认字体，再读回来必然是它，
    /// 与源文件的（Calibri 等）不同——那是既定取舍（见 `CellFormat::is_notable`），
    /// 不是回归。列宽行高也排除：rust_xlsxwriter 会补 ~0.8 字符宽。
    fn snapshot(wb: &Workbook) -> Vec<String> {
        let mut out = vec![format!("active={}", wb.active_sheet)];
        for name in wb.sheet_names() {
            let s = wb.get_sheet(&name).unwrap();
            let mut rows: Vec<String> = s
                .cells()
                .iter()
                .map(|(&(r, c), cell)| {
                    let mut f = cell.format.clone();
                    f.font_size = 0.0;
                    f.font_family = String::new();
                    format!("{name}!{r}:{c}={:?}|{:?}", cell.value, f)
                })
                .collect();
            rows.sort();
            out.extend(rows);
            let mut merges: Vec<_> = s
                .merged_regions()
                .iter()
                .map(|m| format!("{name}!merge {}:{}:{}:{}", m.start_row, m.start_col, m.end_row, m.end_col))
                .collect();
            merges.sort();
            out.extend(merges);
            out.push(format!("{name}!tab={:?}", s.tab_color));
            out.push(format!("{name}!hidden_rows={:?}", {
                let mut v: Vec<_> = s.hidden_rows.iter().copied().collect();
                v.sort();
                v
            }));
            let mut hc: Vec<_> = s.hidden_cols.iter().copied().collect();
            hc.sort();
            out.push(format!("{name}!hidden_cols={hc:?}"));
        }
        out
    }

    #[test]
    fn foreign_xlsx_without_styles_part_is_still_readable() {
        // No styles.xml / theme1.xml at all — the style pass must be a no-op,
        // not a hard error.
        let bytes = build_xlsx(&[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
            ("xl/worksheets/sheet1.xml", E2E_SHEET_XML),
        ]);
        let wb = read_xlsx_from_bytes(&bytes).expect("should still read");
        let sheet = wb.get_sheet("Sheet1").expect("Sheet1");
        assert_eq!(sheet.get_cell(0, 0).expect("A1").format, CellFormat::default());
        // Layout still lands, since it does not depend on styles.xml.
        assert_eq!(sheet.col_widths.get(&1), Some(&24.5));
        assert!(sheet.hidden_rows.contains(&1));
    }

    #[test]
    fn writer_round_trip_preserves_styles_layout_and_merged_title() {
        let mut wb = Workbook::new();
        {
            let sheet = wb.get_sheet_mut("Sheet1").unwrap();
            sheet.set_value(0, 0, CellValue::Text("季度销售".into()));
            sheet.set_value(1, 0, CellValue::Text("北京".into()));
            sheet.set_value(1, 1, CellValue::Number(1234.5));

            let title = sheet.get_cell_mut(0, 0).unwrap();
            title.format.bold = true;
            title.format.font_color = Some("#FFFFFF".into());
            title.format.bg_color = Some("#4472C4".into());
            title.format.h_align = HAlign::Center;

            let money = sheet.get_cell_mut(1, 1).unwrap();
            money.format.number_format = Some("#,##0.00".into());
            money.format.borders.bottom = Some(Border {
                style: BorderStyle::Thin,
                color: "#FF0000".into(),
            });

            sheet.col_widths.insert(1, 20.5);
            sheet.row_heights.insert(0, 28.0);
            sheet.set_tab_color(Some("#00B0F0".into()));
            sheet.hidden_rows.insert(2);
            sheet.hidden_cols.insert(3);
            sheet.merge_cells(0, 0, 0, 2).unwrap();
        }

        let bytes = crate::xlsx_writer::write_xlsx_to_buffer(&wb).expect("write");
        let rt = read_xlsx_from_bytes(&bytes).expect("read back");
        let sheet = rt.get_sheet("Sheet1").expect("Sheet1");

        let title = sheet
            .get_cell(0, 0)
            .expect("合并区左上角的值必须能导出（曾经被 merge_range 清空）");
        assert_eq!(title.value, CellValue::Text("季度销售".into()));
        assert!(title.format.bold);
        assert_eq!(title.format.font_color.as_deref(), Some("#FFFFFF"));
        assert_eq!(title.format.bg_color.as_deref(), Some("#4472C4"));
        assert_eq!(title.format.h_align, HAlign::Center);

        let money = sheet.get_cell(1, 1).expect("B2");
        assert_eq!(money.value, CellValue::Number(1234.5));
        assert_eq!(money.format.number_format.as_deref(), Some("#,##0.00"));
        assert_eq!(
            money.format.borders.bottom.as_ref().map(|b| &b.style),
            Some(&BorderStyle::Thin)
        );

        assert!(
            sheet
                .merged_regions()
                .iter()
                .any(|r| r.start_row == 0 && r.start_col == 0 && r.end_row == 0 && r.end_col == 2),
            "合并区域应被导出并读回"
        );
        // rust_xlsxwriter 会把列宽加约 0.8 字符的 padding（Excel 的字符宽换算法），
        // 因此列宽只做容差比对，不追求逐位相等。
        let width = *sheet.col_widths.get(&1).expect("列宽应回读");
        assert!(
            (width - 20.5).abs() < 1.0,
            "列宽应接近 20.5，实际 {width}"
        );
        let height = *sheet.row_heights.get(&0).expect("行高应回读");
        assert!(
            (height - 28.0).abs() < 0.5,
            "行高应接近 28，实际 {height}（rust_xlsxwriter 按像素换算，会有 1/4 点量化）"
        );
        assert!(sheet.hidden_rows.contains(&2));
        assert!(sheet.hidden_cols.contains(&3));
        assert_eq!(sheet.tab_color.as_deref(), Some("#00B0F0"));
    }

    #[test]
    fn writer_keeps_date_number_format_from_source() {
        let mut wb = Workbook::new();
        {
            let sheet = wb.get_sheet_mut("Sheet1").unwrap();
            sheet.set_value(0, 0, CellValue::Date("2024-03-05".into()));
            sheet.set_value(1, 0, CellValue::Date("2024-03-05".into()));
            sheet.get_cell_mut(0, 0).unwrap().format.number_format =
                Some("yyyy年m月d日".into());
        }
        let bytes = crate::xlsx_writer::write_xlsx_to_buffer(&wb).expect("write");
        let rt = read_xlsx_from_bytes(&bytes).expect("read back");
        let sheet = rt.get_sheet("Sheet1").expect("Sheet1");
        assert_eq!(
            sheet.get_cell(0, 0).unwrap().format.number_format.as_deref(),
            Some("yyyy年m月d日"),
            "源文件的日期格式应被沿用"
        );
        assert_eq!(
            sheet.get_cell(1, 0).unwrap().format.number_format.as_deref(),
            Some("yyyy-mm-dd"),
            "没带格式时回落到默认日期格式"
        );
    }
}
