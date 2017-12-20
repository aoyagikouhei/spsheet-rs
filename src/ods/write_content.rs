use super::{Book,Sheet,Cell,Value};
use super::tempdir::TempDir;
use std::collections::HashMap;
use std::result;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::writer::Writer;
use std::io::Cursor;
use std::io::Write;
use std::fs::File;
use file_common::*;
use super::OdsError;

const CONTENT_XML: &'static str = "content.xml";

fn make_content_xml_none_table_cell(writer: &mut Writer<Cursor<Vec<u8>>>, none_count: i64) {
    if none_count > 0 {
        if none_count == 1 {
            write_start_tag(writer, "table:table-cell", vec![], true);
        } else {
            write_start_tag(writer, "table:table-cell", vec![("table:number-columns-repeated", none_count.to_string().as_str())], true);
        }
    }
}

fn make_content_xml_table_cell(writer: &mut Writer<Cursor<Vec<u8>>>, cell: &Cell, date_hash: &HashMap<String, String>) {
    match cell.get_value() {
        &Value::Str(ref value) => {
            write_start_tag(writer, "table:table-cell", vec![("office:value-type", "string"), ("calcext:value-type", "string")], false);
            write_start_tag(writer, "text:p", vec![], false);
            write_text_node(writer, value.to_string());
        },
        &Value::Float(ref value) => {
            write_start_tag(writer, "table:table-cell", vec![("office:value-type", "float"), ("office:value", value.to_string().as_str()), ("calcext:value-type", "float")], false);
            write_start_tag(writer, "text:p", vec![], false);
            write_text_node(writer, value.to_string());
        },
        &Value::Date(ref value) => {
            let formater = cell.get_style().get_format();
            write_start_tag(writer, "table:table-cell", vec![
                ("table:style-name", date_hash.get(formater).unwrap().as_str()), 
                ("office:value-type", "date"), 
                ("office:date-value", value.format("%Y-%m-%dT%H:%M:%S").to_string().as_str()), 
                ("calcext:value-type", "date")
            ], false);
            write_start_tag(writer, "text:p", vec![], false);
            write_text_node(writer, cell.get_formated_date().unwrap());
        },
    }
    write_end_tag(writer, "text:p");
    write_end_tag(writer, "table:table-cell");
}

fn make_content_xml_by_sheet(writer: &mut Writer<Cursor<Vec<u8>>>, sheet: &Sheet, date_hash: &HashMap<String, String>) {
    write_start_tag(writer, "table:table", vec![("table:name", sheet.get_name().as_str()),("table:style-name", "ta1"),], false);

    let indexes = sheet.get_max_index();
    match indexes {
        Some(indexes) => {
            write_start_tag(writer, "table:table-column", vec![
                ("table:style-name", "co1"),
                ("table:number-columns-repeated", (indexes.1 + 1).to_string().as_str()),
                ("table:default-cell-style-name", "Default")
            ], true);
        },
        None => {
            write_start_tag(writer, "table:table-column", vec![
                ("table:style-name", "co1"),
                ("table:default-cell-style-name", "Default")
            ], true);
        }
    };

    match indexes {
        Some(indexes) => {
            for row_index in 0..indexes.0 + 1 {
                match sheet.get_rows().get(&row_index) {
                    Some(columns) => {
                        write_start_tag(writer, "table:table-row", vec![("table:style-name", "ro1"),], false);
                        let mut max_column_index = 0;
                        for (column_index, _) in columns.iter() {
                            if *column_index > max_column_index {
                                max_column_index = *column_index;
                            }
                        }
                        let mut none_count = 0;
                        for column_index in 0..max_column_index+1 {
                            match columns.get(&column_index) {
                                Some(cell) => {
                                    let _ = make_content_xml_none_table_cell(writer, none_count);
                                    none_count = 0;
                                    let _ = make_content_xml_table_cell(writer, cell, date_hash);
                                },
                                None => {
                                    none_count = none_count + 1;
                                }
                            }
                        }
                        let _ = make_content_xml_none_table_cell(writer, none_count);
                        write_end_tag(writer, "table:table-row");
                    },
                    None => {
                        // some row not found
                        write_start_tag(writer, "table:table-row", vec![
                            ("table:style-name", "ro1")
                        ], false);
                        write_start_tag(writer, "table:table-cell", vec![
                            ("table:number-columns-repeated", (indexes.1 + 1).to_string().as_str())
                        ], true);
                        write_end_tag(writer, "table:table-row");
                    }
                }
            }
        },
        None => {
            // all row not found
            write_start_tag(writer, "table:table-row", vec![
                ("table:style-name", "ro1")
            ], false);
            write_start_tag(writer, "table:table-cell", vec![], true);
            write_end_tag(writer, "table:table-row");
        }
    };
    write_end_tag(writer, "table:table");
}

fn make_number_format(writer: &mut Writer<Cursor<Vec<u8>>>, formats: &Vec<&str>) {
    for it in formats {
        match *it {
            "%Y" => {
                 write_start_tag(writer, "number:year", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%y" => {
                 write_start_tag(writer, "number:year", vec![
                ], true);
            },
            "%m" => {
                 write_start_tag(writer, "number:month", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%-m" => {
                 write_start_tag(writer, "number:month", vec![
                ], true);
            },
            "%d" => {
                 write_start_tag(writer, "number:day", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%-d" => {
                 write_start_tag(writer, "number:day", vec![
                ], true);
            },
            "%H" => {
                 write_start_tag(writer, "number:hours", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%-H" => {
                 write_start_tag(writer, "number:hours", vec![
                ], true);
            },
            "%M" => {
                 write_start_tag(writer, "number:minutes", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%-M" => {
                 write_start_tag(writer, "number:minutes", vec![
                ], true);
            },
            "%S" => {
                 write_start_tag(writer, "number:seconds", vec![
                    ("number:style", "long"),
                ], true);
            },
            "%-S" => {
                 write_start_tag(writer, "number:seconds", vec![
                ], true);
            },
            _ => {
                write_start_tag(writer, "number:text", vec![], false);
                write_text_node(writer, String::from(*it));
                write_end_tag(writer, "number:text");
            }
        }
    }
    
    /*
    let mut percent_flag = false;


    for ch in format.chars() {
        match ch {
            '%' => {},
            'Y' if percent_flag => {
                write_start_tag(writer, "number:year", vec![
                    ("number:style", "long"),
                    ], true);
            },
            'y' if percent_flag => {
                write_start_tag(writer, "number:year", vec![], true);
            },
            'm' if percent_flag => {
                write_start_tag(writer, "number:month", vec![], true);
            },
            'd' if percent_flag => {
                write_start_tag(writer, "number:day", vec![], true);
            },
            'H' if percent_flag => {
                write_start_tag(writer, "number:hours", vec![], true);
            },
            'M' if percent_flag => {
                write_start_tag(writer, "number:minutes", vec![
                    ("number:style", "long"),
                    ], true);
            },
            'S' if percent_flag => {
                write_start_tag(writer, "number:seconds", vec![
                    ("number:style", "long"),
                    ], true);
            },
            _ => {
                write_start_tag(writer, "number:text", vec![], false);
                let mut val = String::from("");
                val.push(ch);
                write_text_node(writer, val);
                write_end_tag(writer, "number:text");
            }
        }
        percent_flag = ch == '%';
    }
    */
}

fn make_num_styles(writer: &mut Writer<Cursor<Vec<u8>>>, book: &Book) -> HashMap<String, String> {
    let mut result = HashMap::new();
    let mut count: usize = 0;
    for sheet in book.get_sheet_vec() {
        sheet.walk_through(|_, _, cell| {
            match cell.get_value() {
                &Value::Date(_) => {
                    let format = cell.get_style().get_format();
                    if !result.contains_key(format) {
                        count = count + 1;
                        let n_name = format!("N{}", count);
                        let s_name = format!("ce{}", count);
                        result.insert(format.clone(), s_name.clone());
                         write_start_tag(writer, "number:date-style", vec![
                            ("style:name", n_name.as_str()),
                            ("number:automatic-order", "true"),
                            ], false);
                        make_number_format(writer, &cell.get_style().get_date_formats().unwrap());
                        write_end_tag(writer, "number:date-style");
                        write_start_tag(writer, "style:style", vec![
                            ("style:name", s_name.as_str()),
                            ("style:family", "table-cell"),
                            ("style:parent-style-name", "Default"),
                            ("style:data-style-name", n_name.as_str()),
                            ], true);
                    }
                },
                _ => {}
            }
        });
    }
    result
}

pub fn write(book: &Book, dir: &TempDir) -> result::Result<(), OdsError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "office:document-content", vec![("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"),("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"),("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"),("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"),("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"),("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),("xmlns:xlink", "http://www.w3.org/1999/xlink"),("xmlns:dc", "http://purl.org/dc/elements/1.1/"),("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"),("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"),("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"),("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"),("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"),("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"),("xmlns:math", "http://www.w3.org/1998/Math/MathML"),("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"),("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"),("xmlns:ooo", "http://openoffice.org/2004/office"),("xmlns:ooow", "http://openoffice.org/2004/writer"),("xmlns:oooc", "http://openoffice.org/2004/calc"),("xmlns:dom", "http://www.w3.org/2001/xml-events"),("xmlns:xforms", "http://www.w3.org/2002/xforms"),("xmlns:xsd", "http://www.w3.org/2001/XMLSchema"),("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"),("xmlns:rpt", "http://openoffice.org/2005/report"),("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"),("xmlns:xhtml", "http://www.w3.org/1999/xhtml"),("xmlns:grddl", "http://www.w3.org/2003/g/data-view#"),("xmlns:tableooo", "http://openoffice.org/2009/table"),("xmlns:drawooo", "http://openoffice.org/2010/draw"),("xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"),("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"),("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"),("xmlns:formx", "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"),("xmlns:css3t", "http://www.w3.org/TR/css3-text/"),("office:version", "1.2"),], false);
    write_start_tag(&mut writer, "office:scripts", vec![], false);
    write_end_tag(&mut writer, "office:scripts");
    write_start_tag(&mut writer, "office:font-face-decls", vec![], false);
    write_start_tag(&mut writer, "style:font-face", vec![("style:name", "Liberation Sans"),("svg:font-family", "&apos;Liberation Sans&apos;"),("style:font-family-generic", "swiss"),("style:font-pitch", "variable"),], false);
    write_end_tag(&mut writer, "style:font-face");
    write_start_tag(&mut writer, "style:font-face", vec![("style:name", "Arial Unicode MS"),("svg:font-family", "&apos;Arial Unicode MS&apos;"),("style:font-family-generic", "system"),("style:font-pitch", "variable"),], false);
    write_end_tag(&mut writer, "style:font-face");
    write_start_tag(&mut writer, "style:font-face", vec![("style:name", "Tahoma"),("svg:font-family", "Tahoma"),("style:font-family-generic", "system"),("style:font-pitch", "variable"),], false);
    write_end_tag(&mut writer, "style:font-face");
    write_start_tag(&mut writer, "style:font-face", vec![("style:name", "ヒラギノ明朝 ProN"),("svg:font-family", "&apos;ヒラギノ明朝 ProN&apos;"),("style:font-family-generic", "system"),("style:font-pitch", "variable"),], false);
    write_end_tag(&mut writer, "style:font-face");
    write_end_tag(&mut writer, "office:font-face-decls");
    write_start_tag(&mut writer, "office:automatic-styles", vec![], false);
    write_start_tag(&mut writer, "style:style", vec![("style:name", "co1"),("style:family", "table-column"),], false);
    write_start_tag(&mut writer, "style:table-column-properties", vec![("fo:break-before", "auto"),("style:column-width", "22.58mm"),], false);
    write_end_tag(&mut writer, "style:table-column-properties");
    write_end_tag(&mut writer, "style:style");
    write_start_tag(&mut writer, "style:style", vec![("style:name", "ro1"),("style:family", "table-row"),], false);
    write_start_tag(&mut writer, "style:table-row-properties", vec![("style:row-height", "4.52mm"),("fo:break-before", "auto"),("style:use-optimal-row-height", "true"),], false);
    write_end_tag(&mut writer, "style:table-row-properties");
    write_end_tag(&mut writer, "style:style");
    write_start_tag(&mut writer, "style:style", vec![("style:name", "ta1"),("style:family", "table"),("style:master-page-name", "Default"),], false);
    write_start_tag(&mut writer, "style:table-properties", vec![("table:display", "true"),("style:writing-mode", "lr-tb"),], false);
    write_end_tag(&mut writer, "style:table-properties");
    write_end_tag(&mut writer, "style:style");

    let date_hash = make_num_styles(&mut writer, book);

    write_end_tag(&mut writer, "office:automatic-styles");
    write_start_tag(&mut writer, "office:body", vec![], false);
    write_start_tag(&mut writer, "office:spreadsheet", vec![], false);
    write_start_tag(&mut writer, "table:calculation-settings", vec![("table:automatic-find-labels", "false"),("table:use-regular-expressions", "false"),("table:use-wildcards", "true"),], false);
    write_end_tag(&mut writer, "table:calculation-settings");

    for sheet in book.get_sheet_vec() {
        let _ = make_content_xml_by_sheet(&mut writer, &sheet, &date_hash);
    }

    write_start_tag(&mut writer, "table:named-expressions", vec![], false);
    write_end_tag(&mut writer, "table:named-expressions");
    write_end_tag(&mut writer, "office:spreadsheet");
    write_end_tag(&mut writer, "office:body");
    write_end_tag(&mut writer, "office:document-content");

    let file_path = dir.path().join(CONTENT_XML);
    let mut f = File::create(file_path)?;
    f.write_all(writer.into_inner().get_ref())?;
    f.sync_all()?;

    Ok(())
}