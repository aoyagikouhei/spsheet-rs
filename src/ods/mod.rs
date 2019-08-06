//! OpenDocument ods read and write
extern crate quick_xml;
extern crate tempdir;
extern crate zip;

use self::quick_xml::events::{BytesStart, Event};
use self::quick_xml::Reader;
use self::tempdir::TempDir;
use super::{Book, Cell, Sheet, Value};
use file_common::*;
use std::fs::File;
use std::io;
use std::io::BufReader;
use std::path::Path;
use std::result;
use std::string::FromUtf8Error;

mod read_content;
mod read_style;
mod write_content;
mod write_style;

const MANIFEST_XML_CONTENT: &'static str = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
 <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
 <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
 <manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#;

#[derive(Debug)]
pub enum OdsError {
    Io(io::Error),
    Xml(quick_xml::Error),
    Zip(zip::result::ZipError),
    Uft8(FromUtf8Error),
}

impl From<io::Error> for OdsError {
    fn from(err: io::Error) -> OdsError {
        OdsError::Io(err)
    }
}

impl From<quick_xml::Error> for OdsError {
    fn from(err: quick_xml::Error) -> OdsError {
        OdsError::Xml(err)
    }
}

impl From<zip::result::ZipError> for OdsError {
    fn from(err: zip::result::ZipError) -> OdsError {
        OdsError::Zip(err)
    }
}

impl From<FromUtf8Error> for OdsError {
    fn from(err: FromUtf8Error) -> OdsError {
        OdsError::Uft8(err)
    }
}

type Result<T> = result::Result<T, OdsError>;

pub fn read(path: &Path) -> Result<Book> {
    let file = File::open(path)?;
    let dir = TempDir::new("shreadsheet")?;
    match unzip(&file, &dir) {
        Ok(_) => {}
        Err(err) => {
            dir.close()?;
            return Err(OdsError::Zip(err));
        }
    }
    let style_content = read_style::read(&dir).unwrap();
    let book = read_content::read(&dir, &style_content);
    dir.close()?;
    book
}

pub fn write(book: &Book, path: &Path) -> result::Result<(), OdsError> {
    let dir = TempDir::new("shreadsheet")?;
    let _ = write_style::write(book, &dir);
    let _ = write_content::write(book, &dir);
    let _ = make_static_file(
        &dir,
        "META-INF/manifest.xml",
        MANIFEST_XML_CONTENT,
        Some("META-INF"),
    )?;
    write_to_file(path, &dir)?;
    dir.close()?;
    Ok(())
}

fn read_number_format(
    e: &BytesStart,
    long_value: &str,
    short_value: &str,
) -> result::Result<String, OdsError> {
    let mut number_style = String::from("");
    for a in e.attributes().with_checks(false) {
        match a {
            Ok(ref attr) if attr.key == b"number:style" => {
                number_style = get_attribute_value(attr)?;
            }
            Ok(_) => {}
            Err(_) => {}
        }
    }
    Ok(String::from(if number_style == "long" {
        long_value
    } else {
        short_value
    }))
}

fn read_number_year(e: &BytesStart) -> result::Result<String, OdsError> {
    let mut number_style = String::from("");
    let mut number_calendar = String::from("");
    for a in e.attributes().with_checks(false) {
        match a {
            Ok(ref attr) if attr.key == b"number:style" => {
                number_style = get_attribute_value(attr)?;
            }
            Ok(ref attr) if attr.key == b"number:calendar" => {
                number_calendar = get_attribute_value(attr)?;
            }
            Ok(_) => {}
            Err(_) => {}
        }
    }
    Ok(String::from(
        if number_style == "long" && number_calendar == "gengou" {
            "EE"
        } else if number_calendar == "gengou" {
            "E"
        } else if number_style == "long" {
            "YYYY"
        } else {
            "YY"
        },
    ))
}

fn read_number_date_style(
    reader: &mut Reader<BufReader<File>>,
) -> result::Result<String, OdsError> {
    let mut buf = Vec::new();
    let mut style_format = String::from("");
    let mut text_empty_flag = true;
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name() {
                b"number:text" => {
                    text_empty_flag = true;
                }
                _ => (),
            },
            Ok(Event::End(ref e)) => match e.name() {
                b"number:text" => {
                    if text_empty_flag {
                        style_format.push_str("\\ ");
                    }
                }
                b"number:date-style" => {
                    return Ok(style_format);
                }
                _ => (),
            },
            Ok(Event::Empty(ref e)) => {
                let added_string = match e.name() {
                    b"number:era" => read_number_format(e, "GGG", "G"),
                    b"number:year" => read_number_year(e),
                    b"number:month" => read_number_format(e, "MM", "M"),
                    b"number:day" => read_number_format(e, "DD", "D"),
                    b"number:hours" => read_number_format(e, "HH", "H"),
                    b"number:minutes" => read_number_format(e, "MM", "M"),
                    b"number:seconds" => read_number_format(e, "SS", "S"),
                    _ => Ok(String::from("")),
                };
                style_format.push_str(added_string?.as_str());
            }
            Ok(Event::Text(e)) => {
                match e.unescape_and_decode(&reader).unwrap().as_str() {
                    "/" => style_format.push_str("/"),
                    ":" => style_format.push_str(":"),
                    other => {
                        if other.chars().count() == 1 {
                            style_format.push_str("\\");
                            style_format.push_str(other);
                        } else {
                            style_format.push_str("\"");
                            style_format.push_str(other);
                            style_format.push_str("\"");
                        }
                    }
                }
                text_empty_flag = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }

        // if we don't keep a borrow elsewhere, we can clear the buffer to keep memory usage low
        buf.clear();
    }
    Ok(String::from(""))
}
