use file_common::*;
use std::io::Cursor;
use std::result;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::writer::Writer;
use super::tempdir::TempDir;
use super::{Book};
use super::XlsxError;

const WORKBOOK_XML_RELS: &'static str = "xl/_rels/workbook.xml.rels";

pub fn write(book: &Book, dir: &TempDir) -> result::Result<(), XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(
        BytesDecl::new(b"1.0", Some(b"UTF-8"), None)));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "Relationships", vec![
        ("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
    ], false);
    let size = book.get_sheet_size();
    for i in 0..size {
        let index = (i+1).to_string();
        write_start_tag(&mut writer, "Relationship", vec![
            ("Id", format!("rId{}", index).as_str()),
            ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
            ("Target", format!("worksheets/sheet{}.xml", index).as_str())
        ], true);
    }
    write_start_tag(&mut writer, "Relationship", vec![
        ("Id", format!("rId{}", size + 1).as_str()),
        ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
        ("Target", "styles.xml")
    ], true);
    write_start_tag(&mut writer, "Relationship", vec![
        ("Id", format!("rId{}", size + 2).as_str()),
        ("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"),
        ("Target", "sharedStrings.xml")
    ], true);
    write_end_tag(&mut writer, "Relationships");
    let _ = make_file_from_writer(WORKBOOK_XML_RELS, dir, writer, Some("xl/_rels"))?;
    Ok(())
}