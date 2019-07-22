use file_common::*;
use std::collections::HashMap;
use std::io::Cursor;
use std::result;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::Writer;
use super::tempdir::TempDir;
use super::{Book, Value};
use super::XlsxError;

const STYLE_XML: &'static str = "xl/styles.xml";

fn make_num_fmts(writer: &mut Writer<Cursor<Vec<u8>>>, book: &Book) -> Vec<HashMap<String, String>> {
    let mut result = vec![];
    let mut num_fmot_id = 164;
    let mut key_map = HashMap::new();
    for sheet in book.get_sheet_vec() {
        sheet.walk_through(|_, _, cell| {
            match cell.get_value() {
                &Value::Date(_) => {
                    let format = cell.get_format().get_content();
                    if !key_map.contains_key(format) {
                        let mut map = HashMap::new();
                        map.insert(String::from("numFmtId"), num_fmot_id.to_string());
                        map.insert(String::from("format"), format.clone());
                        result.push(map);
                        num_fmot_id = num_fmot_id + 1;
                        key_map.insert(format.clone(), ());
                    }
                },
                _ => {}
            }
        });
    }

    write_start_tag(writer, "numFmts", vec![
        ("count", result.len().to_string().as_str()),
    ], false);

    for map in result.iter() {
        let num_fmt_id = map.get(&String::from("numFmtId")).unwrap();
        let format_code = map.get(&String::from("format")).unwrap();
        write_start_tag(writer, "numFmt", vec![
            ("numFmtId", num_fmt_id.as_str()),
            ("formatCode", format_code.as_str()),
        ], true);
    }

    write_end_tag(writer, "numFmts");
    result
}

fn make_cell_xfs(writer: &mut Writer<Cursor<Vec<u8>>>, num_fmts: &Vec<HashMap<String, String>>) -> HashMap<String, usize> {
    let mut result: HashMap<String, usize> = HashMap::new();
    let count = num_fmts.len() + 1;
    write_start_tag(writer, "cellXfs", vec![("count", count.to_string().as_str()),], false);
    write_start_tag(writer, "xf", vec![
        ("borderId", "0"),
        ("fillId", "0"),
        ("fontId", "0"),
        ("numFmtId", "0"),
        ("xfId", "0"),
        ("applyAlignment", "1"),
        ("applyFont", "1"),
    ], false);
    write_start_tag(writer, "alignment", vec![
        ("readingOrder", "0"),
        ("shrinkToFit", "0"),
        ("vertical", "bottom"),
        ("wrapText", "0"),
    ], true);
    write_end_tag(writer, "xf");

    let mut count = 0;
    for map in num_fmts.iter() {
        count = count + 1;
        let num_fmt_id = map.get(&String::from("numFmtId")).unwrap();
        let format_code = map.get(&String::from("format")).unwrap();
        result.insert(format_code.clone(), count as usize);
        write_start_tag(writer, "xf", vec![
            ("borderId", "0"),
            ("fillId", "0"),
            ("fontId", "1"),
            ("numFmtId", num_fmt_id.to_string().as_str()),
            ("xfId", "0"),
            ("applyAlignment", "1"),
            ("applyFont", "1"),
        ], false);
        write_start_tag(writer, "alignment", vec![
            ("readingOrder", "0"),
        ], true);
        write_end_tag(writer, "xf");
    }

    write_end_tag(writer, "cellXfs");
    result
}

pub fn write(book: &Book, dir: &TempDir) -> result::Result<HashMap<String, usize>, XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(
        BytesDecl::new(b"1.0", Some(b"UTF-8"), Some(b"yes"))));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "styleSheet", vec![("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"),("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),], false);

    let num_fmts = make_num_fmts(&mut writer, book);

    write_start_tag(&mut writer, "fonts", vec![("count", "2"),], false);
    write_start_tag(&mut writer, "font", vec![], false);
    write_start_tag(&mut writer, "sz", vec![("val", "10.0"),], false);
    write_end_tag(&mut writer, "sz");
    write_start_tag(&mut writer, "color", vec![("rgb", "FF000000"),], false);
    write_end_tag(&mut writer, "color");
    write_start_tag(&mut writer, "name", vec![("val", "Arial"),], false);
    write_end_tag(&mut writer, "name");
    write_end_tag(&mut writer, "font");
    write_start_tag(&mut writer, "font", vec![], false);
    write_end_tag(&mut writer, "font");
    write_end_tag(&mut writer, "fonts");
    write_start_tag(&mut writer, "fills", vec![("count", "2"),], false);
    write_start_tag(&mut writer, "fill", vec![], false);
    write_start_tag(&mut writer, "patternFill", vec![("patternType", "none"),], false);
    write_end_tag(&mut writer, "patternFill");
    write_end_tag(&mut writer, "fill");
    write_start_tag(&mut writer, "fill", vec![], false);
    write_start_tag(&mut writer, "patternFill", vec![("patternType", "lightGray"),], false);
    write_end_tag(&mut writer, "patternFill");
    write_end_tag(&mut writer, "fill");
    write_end_tag(&mut writer, "fills");
    write_start_tag(&mut writer, "borders", vec![("count", "1"),], false);
    write_start_tag(&mut writer, "border", vec![], false);
    write_end_tag(&mut writer, "border");
    write_end_tag(&mut writer, "borders");
    write_start_tag(&mut writer, "cellStyleXfs", vec![("count", "1"),], false);
    write_start_tag(&mut writer, "xf", vec![("borderId", "0"),("fillId", "0"),("fontId", "0"),("numFmtId", "0"),("applyAlignment", "1"),("applyFont", "1"),], false);
    write_end_tag(&mut writer, "xf");
    write_end_tag(&mut writer, "cellStyleXfs");

    let result = make_cell_xfs(&mut writer, &num_fmts);

    write_start_tag(&mut writer, "cellStyles", vec![("count", "1"),], false);
    write_start_tag(&mut writer, "cellStyle", vec![("xfId", "0"),("name", "Normal"),("builtinId", "0"),], false);
    write_end_tag(&mut writer, "cellStyle");
    write_end_tag(&mut writer, "cellStyles");
    write_start_tag(&mut writer, "dxfs", vec![("count", "0"),], false);
    write_end_tag(&mut writer, "dxfs");
    write_end_tag(&mut writer, "styleSheet");

    let _ = make_file_from_writer(STYLE_XML, dir, writer, Some("xl"))?;
    Ok(result)
}
