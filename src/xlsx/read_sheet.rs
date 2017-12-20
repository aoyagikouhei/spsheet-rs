use file_common::*;
use std::collections::HashMap;
use std::result;
use super::time::Duration;
use super::chrono::prelude::*;
use super::quick_xml::reader::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use super::XlsxError;
use super::{Sheet,Cell,Style,Value,column_and_row_to_index};

pub fn read(dir: &TempDir, name: &String, target: &String, shared_strings: &Vec<String>, styles: &Vec<HashMap<String, String>>) -> result::Result<Sheet, XlsxError> {
    let mut sheet = Sheet::new(name.as_str());

    let path = dir.path().join("xl/".to_string() + target);
    let mut reader = Reader::from_file(path)?;
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut row_index: usize = 0;
    let mut column_index: usize = 0;
    let mut string_value: String = String::from("");
    let mut type_value: String = String::from("");
    let mut style_index: usize = 0;

    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"row" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"r" => {
                                    let value = get_attribute_value(attr)?;
                                    row_index = value.parse::<usize>().unwrap() - 1;
                                },
                                Ok(_) => {},
                                Err(_) => {},
                            }
                        }
                    },
                    b"c" => {
                        for a in e.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"r" => {
                                    let value = get_attribute_value(attr)?;
                                    // A3のような値からcolumn_indexを計算する
                                    column_index = column_and_row_to_index(value).unwrap().0;
                                },
                                Ok(ref attr) if attr.key == b"s" => {
                                    let value = get_attribute_value(attr)?;
                                    style_index = value.parse::<usize>().unwrap();
                                },
                                Ok(ref attr) if attr.key == b"t" => {
                                    type_value = get_attribute_value(attr)?;
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
                    b"v" => {
                        let cell = if type_value == "s" {
                            let index = string_value.parse::<usize>().unwrap();
                            let val = shared_strings.get(index).unwrap();
                            Cell::str((*val).clone())
                        } else {
                            let hash = &styles[style_index];
                            match hash.get("formatCode") {
                                Some(format_code) => {
                                    Cell::new(
                                        Value::Date(number_to_date(&string_value)), 
                                        Style::new(format_code.to_string())
                                    )
                                },
                                None => {
                                    Cell::float(string_value.parse::<f64>().unwrap())
                                }
                            }
                        };
                        sheet.add_cell(cell, row_index, column_index);
                    },
                    _ => (),
                }
            },
            Ok(Event::Text(e)) => string_value = e.unescape_and_decode(&reader).unwrap(),
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }
        buf.clear();
    }

    Ok(sheet)
}

// 1900年からのepoch
// 43071.5625 -> 2017-12-02T13:30:00
fn number_to_date(src: &String) -> DateTime<Utc> {
    let num = src.parse::<f64>().unwrap();
    let timestamp = ((((num - num.floor()) * 86400.0) as f64).round()) as i64;
    let hms = NaiveDateTime::from_timestamp(timestamp, 0);

    let dt = (Utc.ymd(1900, 1, 1) + Duration::days(num.floor() as i64 - 2)).and_hms(hms.hour(), hms.minute(), hms.second());
    dt
}