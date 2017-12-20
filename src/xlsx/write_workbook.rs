use file_common::*;
use std::io::Cursor;
use std::result;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::writer::Writer;
use super::tempdir::TempDir;
use super::{Book};
use super::XlsxError;

const WORKBOOK_XML: &'static str = "xl/workbook.xml";

pub fn write(book: &Book, dir: &TempDir) -> result::Result<(), XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(
        BytesDecl::new(b"1.0", Some(b"UTF-8"), Some(b"yes"))));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "workbook", vec![
        ("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
        ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ], false);
    write_start_tag(&mut writer, "fileVersion", vec![
        ("appName", "Calc")
    ], true);
    write_start_tag(&mut writer, "workbookPr", vec![
        ("backupFile", "false"),
        ("showObjects", "all"),
        ("date1904", "false")
    ], true);
    write_start_tag(&mut writer, "workbookProtection", vec![
    ], true);
    write_start_tag(&mut writer, "bookViews", vec![
    ], false);
    write_start_tag(&mut writer, "workbookView", vec![
        ("showHorizontalScroll", "true"),
        ("showVerticalScroll", "true"),
        ("showSheetTabs", "true"),
        ("xWindow", "0"),
        ("yWindow", "0"),
        ("windowWidth", "16384"),
        ("windowHeight", "8192"),
        ("tabRatio", "500"),
        ("firstSheet", "0"),
        ("activeTab", "0")
    ], true);
    write_end_tag(&mut writer, "bookViews");
    write_start_tag(&mut writer, "sheets", vec![], false);
    let mut index = 1;
    for sheet in book.get_sheet_vec() {
        write_start_tag(&mut writer, "sheet", vec![
            ("name", sheet.get_name()),
            ("sheetId", index.to_string().as_str()),
            ("state", "visible"),
            ("r:id", format!("rId{}", index.to_string()).as_str())
        ], true);
        index = index + 1;
    }
    write_end_tag(&mut writer, "sheets");
    write_start_tag(&mut writer, "calcPr", vec![
        ("iterateCount", "100"),
        ("refMode", "A1"),
        ("iterate", "false"),
        ("iterateDelta", "0.001")
    ], true);
    write_start_tag(&mut writer, "extLst", vec![
    ], false);
    write_start_tag(&mut writer, "ext", vec![
        ("xmlns:loext", "http://schemas.libreoffice.org/"),
        ("uri", "{7626C862-2A13-11E5-B345-FEFF819CDC9F}")
    ], false);
    write_start_tag(&mut writer, "loext:extCalcPr", vec![
        ("stringRefSyntax", "CalcA1ExcelA1")
    ], true);
    write_end_tag(&mut writer, "ext");
    write_end_tag(&mut writer, "extLst");
    write_end_tag(&mut writer, "workbook");
    let _ = make_file_from_writer(WORKBOOK_XML, dir, writer, Some("xl"))?;
    Ok(())
}