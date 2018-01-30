use super::{Book,Sheet,Cell};
use file_common::*;
use super::quick_xml::reader::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use super::{Result};
use super::read_style::StyleContent;
use std::collections::HashMap;

const CONTENT_XML: &'static str = "content.xml";

pub fn read(dir: &TempDir, style_content: &StyleContent) -> Result<Book> {
    let mut date_style_map = HashMap::new();
    let mut style_map_for_date: HashMap<String, String> = HashMap::new();

    let path = dir.path().join(CONTENT_XML);
    let mut reader = Reader::from_file(path)?;
    reader.trim_text(true);
    let mut book = Book::new();

    let mut buf = Vec::new();

    let mut sheet = Sheet::new("");
    let mut row: usize = 0;
    let mut column: usize = 0;
    let mut cell_type: String = String::from("");
    let mut float_value: f64 = 0.0;
    let mut str_value: String = String::from("");
    let mut date_value: String = String::from("");
    let mut table_style_name: String = String::from("");

    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"table:table" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"table:name" => {
                                    sheet.set_name(get_attribute_value(attr)?);
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                    },
                    b"table:table-cell" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"table:style-name" => {
                                    table_style_name = get_attribute_value(attr)?;
                                },
                                Ok(ref attr) if attr.key == b"office:value-type" => {
                                    cell_type = get_attribute_value(attr)?;
                                },
                                Ok(ref attr) if attr.key == b"office:value" => {
                                    let value = get_attribute_value(attr)?;
                                    float_value = value.parse::<f64>().unwrap();
                                },
                                Ok(ref attr) if attr.key == b"office:date-value" => {
                                    date_value = get_attribute_value(attr)?;
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                    },
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
            Ok(Event::End(ref e)) => {
                match e.name() {
                    b"table:table" => {
                        row = 0;
                        column = 0;
                        book.add_sheet(sheet);
                        sheet = Sheet::new("");
                    },
                    b"table:table-row" => {
                        row = row + 1;
                        column = 0;
                    },
                    b"table:table-cell" => {
                        match cell_type.as_str() {
                            "string" => {
                                let cell = Cell::str(str_value.clone(), String::from(""));
                                sheet.add_cell(cell, row, column);
                            },
                            "float" => {
                                let cell = Cell::float(float_value, "");
                                sheet.add_cell(cell, row, column);
                            },
                            "date" => {
                                let format = match style_map_for_date.get(&table_style_name) {
                                    Some(value) => value.clone(),
                                    None => String::from(""),
                                };
                                let cell = Cell::date(date_value.clone(), format);
                                sheet.add_cell(cell, row, column);
                            },
                            _ => {},
                        }
                        cell_type = String::from("");
                        column = column + 1;
                    },
                    _ => (),
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.name() {
                    b"table:table-cell" => {
                        if 0 == e.attributes().count() {
                            column = column + 1;
                        } else {
                            for a in e.attributes().with_checks(false) {
                                match a {
                                    Ok(ref attr) if attr.key == b"table:number-columns-repeated" => {
                                        let value = get_attribute_value(attr)?;
                                        column = column + value.parse::<usize>().unwrap();
                                    },
                                    Ok(_) => {},
                                    Err(_) => {},
                                }
                            }
                        }
                    },
                    b"style:style" => {
                        let mut style_name = String::from("");
                        let mut data_style_name = String::from("");
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"style:name" => {
                                   style_name = get_attribute_value(attr)?;
                                },
                                Ok(ref attr) if attr.key == b"style:data-style-name" => {
                                   data_style_name = get_attribute_value(attr)?;
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                        if data_style_name != "" {
                            if let Some(format) = date_style_map.get(&data_style_name) {
                                style_map_for_date.insert(style_name.clone(), format.clone());
                            } else if let Some(format) = style_content.date_style_map.get(&data_style_name) {
                                style_map_for_date.insert(style_name.clone(), format.clone());
                            }
                        }
                    },
                    _ => (),
                }
            }
            Ok(Event::Text(e)) => str_value = e.unescape_and_decode(&reader).unwrap(),
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }

        // if we don't keep a borrow elsewhere, we can clear the buffer to keep memory usage low
        buf.clear();
    }

    Ok(book)
}