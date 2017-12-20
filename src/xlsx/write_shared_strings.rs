use file_common::*;
use std::collections::HashMap;
use std::io::Cursor;
use std::result;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::writer::Writer;
use super::tempdir::TempDir;
use super::{Book,Value};
use super::XlsxError;

const SHARED_STRINGS: &'static str = "xl/sharedStrings.xml";

pub fn write(book: &Book, dir: &TempDir) -> result::Result<HashMap<String, usize>, XlsxError> {
    let mut shared_strings: Vec<String> = Vec::new();
    let mut count: usize = 0;
    for sheet in book.get_sheet_vec() {
        sheet.sorted_access(|_, _, cell| {
            match cell.get_value() {
                &Value::Str(ref val) => {
                    count = count + 1;
                    if !shared_strings.contains(val) {
                        shared_strings.push(val.clone());
                    }
                },
                _ => {}
            }
        });
    }
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(
        BytesDecl::new(b"1.0", Some(b"UTF-8"), Some(b"yes"))));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "sst", vec![
        ("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), 
        ("count", count.to_string().as_str()), 
        ("uniqueCount", shared_strings.len().to_string().as_str())], false);
    let mut map: HashMap<String, usize> = HashMap::new();
    let mut index = 0;
    for st in shared_strings {
         write_start_tag(&mut writer, "si", vec![], false);
         write_start_tag(&mut writer, "t", vec![("xml:space", "preserve")], false);
         write_text_node(&mut writer, st.clone());
         write_end_tag(&mut writer, "t");
         write_end_tag(&mut writer, "si");
         map.insert(st, index);
         index = index + 1;
    }
    write_end_tag(&mut writer, "sst");
    let _ = make_file_from_writer(SHARED_STRINGS, dir, writer, Some("xl"))?;
    Ok(map)
}