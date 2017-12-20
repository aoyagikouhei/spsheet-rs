use std::result;
use super::quick_xml::reader::Reader;
use super::quick_xml::events::{Event};
use super::tempdir::TempDir;
use super::XlsxError;

const SHARED_STRINGS: &'static str = "xl/sharedStrings.xml";

pub fn read(dir: &TempDir) -> result::Result<Vec<String>, XlsxError> {
    let path = dir.path().join(SHARED_STRINGS);
    let mut reader = Reader::from_file(path)?;
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut res: Vec<String> = Vec::new();
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Text(e)) => res.push(e.unescape_and_decode(&reader).unwrap()),
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }
        buf.clear();
    }
    Ok(res)
}