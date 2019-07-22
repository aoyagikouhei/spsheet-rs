use file_common::*;
use std::collections::HashMap;
use std::result;
use super::quick_xml::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use super::XlsxError;

const WORKBOOK_XML_RELS: &'static str = "xl/_rels/workbook.xml.rels";

pub fn read(dir: &TempDir) -> result::Result<Vec<HashMap<&str, String>>, XlsxError> {
    let path = dir.path().join(WORKBOOK_XML_RELS);
    let mut reader = Reader::from_file(path)?;
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut res: Vec<HashMap<&str, String>> = Vec::new();
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Empty(ref e)) => {
                match e.name() {
                    b"Relationship" => {
                        let mut map: HashMap<&str, String> = HashMap::new();
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"Id" => {
                                    map.insert("id", get_attribute_value(attr)?);
                                },
                                Ok(ref attr) if attr.key == b"Type" => {
                                    map.insert("type", get_attribute_value(attr)?);
                                },
                                Ok(ref attr) if attr.key == b"Target" => {
                                    map.insert("target", get_attribute_value(attr)?);
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                        res.push(map);
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
    Ok(res)
}
