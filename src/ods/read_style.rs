use file_common::*;
use super::quick_xml::reader::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use std::collections::HashMap;
use super::OdsError;

const STYLES_XML: &'static str = "styles.xml";

#[derive(Debug, Clone, PartialEq)]
pub struct StyleContent {
    pub date_style_map: HashMap<String, String>,
}

pub fn read(dir: &TempDir) -> Result<StyleContent, OdsError> {
    let mut date_style_map = HashMap::new();
    
    let path = dir.path().join(STYLES_XML);
    let mut reader = Reader::from_file(path)?;

    reader.trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"number:date-style" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"style:name" => {
                                    date_style_map.insert(
                                        get_attribute_value(attr)?, 
                                        super::read_number_date_style(&mut reader)?);
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                    },
                    _ => (),
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }

        // if we don't keep a borrow elsewhere, we can clear the buffer to keep memory usage low
        buf.clear();
    }

    Ok(StyleContent {
        date_style_map: date_style_map
    })
}