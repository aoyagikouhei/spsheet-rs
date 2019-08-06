use file_common::*;
use std::collections::HashMap;
use std::result;
use super::quick_xml::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use super::XlsxError;

const STYLE_XML: &'static str = "xl/styles.xml";

pub fn read(dir: &TempDir) -> result::Result<Vec<HashMap<String, String>>, XlsxError> {
    let path = dir.path().join(STYLE_XML);
    let mut reader = Reader::from_file(path)?;
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut cell_xfs_flag = false;
    let mut cell_xfs: Vec<HashMap<String, String>> = Vec::new();
    let mut num_fmts: HashMap<String, String> = HashMap::new();
    let mut num_fmt_id = String::from("");
    let mut format_code = String::from("");
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"cellXfs" => {
                        cell_xfs_flag = true;
                    },
                    b"xf" => {
                        let mut map: HashMap<String, String> = HashMap::new();
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"numFmtId" => {
                                    num_fmt_id = get_attribute_value(attr)?;
                                    match num_fmts.get(&num_fmt_id) {
                                        Some(val) => {
                                            map.insert(String::from("formatCode"), val.clone());
                                        },
                                        None => {},
                                    }
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                        if cell_xfs_flag {
                            cell_xfs.push(map);
                        }
                    },
                    _ => (),
                }
            },
            Ok(Event::Empty(ref e)) => {
                match e.name() {
                    b"numFmt" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"numFmtId" => {
                                    num_fmt_id = get_attribute_value(attr)?;
                                },
                                Ok(ref attr) if attr.key == b"formatCode" => {
                                    format_code = get_attribute_value(attr)?;
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                        num_fmts.insert(num_fmt_id.clone(), condvert_character_reference(&format_code));
                    },
                    _ => (),
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }
        buf.clear();
    }
    Ok(cell_xfs)
}
