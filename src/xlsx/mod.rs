//! Office Open XML xlsx read and write
extern crate chrono;
extern crate quick_xml;
extern crate time;
extern crate tempdir;
extern crate walkdir;
extern crate zip;

use file_common::*;
use std::collections::HashMap;
use std::io;
use std::path::Path;
use std::result;
use std::fs::File;
use std::string::FromUtf8Error;
use self::chrono::prelude::*;
use self::tempdir::TempDir;
use super::{Book,Sheet,Cell,Value,column_and_row_to_index,index_to_column};

mod read_sheet;
mod read_shared_strings;
mod read_styles;
mod read_workbook_xml_rels;
mod read_workbook;
mod write_sheet;
mod write_shared_strings;
mod write_styles;
mod write_workbook;
mod write_workbook_xml_rels;

const CONTENT_TYPE_XML: &'static str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;
const APP_XML: &'static str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template></Template><TotalTime>11</TotalTime><Application>spreadsheet-rs/0.0.1</Application></Properties>"#;
const CORE_XML: &'static str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">XXXXXXXXXX</dcterms:created><dc:creator></dc:creator><dc:description></dc:description><dc:language>ja-JP</dc:language><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:modified xsi:type="dcterms:W3CDTF">XXXXXXXXXX</dcterms:modified><cp:revision>6</cp:revision><dc:subject></dc:subject><dc:title></dc:title></cp:coreProperties>"#;
const RELS: &'static str = "_rels/.rels";
const RELS_CONTENT: &'static str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;

#[derive(Debug)]
pub enum XlsxError {
    Io(io::Error),
    Xml(quick_xml::Error),
    Zip(zip::result::ZipError),
    Uft8(FromUtf8Error),
}

impl From<io::Error> for XlsxError {
    fn from(err: io::Error) -> XlsxError {
        XlsxError::Io(err)
    }
}

impl From<quick_xml::Error> for XlsxError {
    fn from(err: quick_xml::Error) -> XlsxError {
        XlsxError::Xml(err)
    }
}

impl From<zip::result::ZipError> for XlsxError {
    fn from(err: zip::result::ZipError) -> XlsxError {
        XlsxError::Zip(err)
    }
}

impl From<FromUtf8Error> for XlsxError {
    fn from(err: FromUtf8Error) -> XlsxError {
        XlsxError::Uft8(err)
    }
}

type Result<T> = result::Result<T, XlsxError>;

pub fn read(path: &Path) -> Result<Book> {
    let file = File::open(path)?;
    let dir = TempDir::new("shreadsheet")?;
    match unzip(&file, &dir) {
        Ok(_) => {},
        Err(err) => {
            dir.close()?;
            return Err(XlsxError::Zip(err));
        }
    }
    let mut book = Book::new();
    {
        let styles = read_styles::read(&dir)?;
        let rels = read_workbook_xml_rels::read(&dir)?;
        let mut rels_map = HashMap::new();
        for r in &rels {
            rels_map.insert(r.get("id").unwrap(), r.get("target").unwrap());
        }
        let sheets = read_workbook::read(&dir)?;
        let shared_strings = read_shared_strings::read(&dir)?;
        for s in &sheets {
            let sheet_target = rels_map.get(s.get("rid").unwrap()).unwrap();
            book.add_sheet(
                read_sheet::read(
                    &dir, s.get("name").unwrap(),
                    sheet_target,
                    &shared_strings,
                    &styles)?);
        }
    }
    dir.close()?;
    Ok(book)
}

pub fn write(book: &Book, path: &Path) -> result::Result<(), XlsxError> {
    let dir = TempDir::new("shreadsheet")?;
    let now = Utc::now();
    let now_str = now.format("%Y-%m-%dT%H:%M:%SZ").to_string();
    let _ = make_static_file(
        &dir, RELS,
        RELS_CONTENT,
        Some("_rels"))?;
    let _ = make_static_file(
        &dir, "[Content_Types].xml",
        CONTENT_TYPE_XML,
        None)?;
    let _ = make_static_file(
        &dir, "docProps/app.xml",
        APP_XML,
        Some("docProps"))?;
    let _ = make_static_file(
        &dir, "docProps/core.xml",
        CORE_XML.replace("XXXXXXXXXX", now_str.as_str()).as_str(),
        Some("docProps"))?;
    let format_map = write_styles::write(book, &dir)?;
    let shared_strings = write_shared_strings::write(book, &dir)?;
    let _ = write_workbook_xml_rels::write(book, &dir)?;
    let _ = write_workbook::write(book, &dir)?;
    let mut index = 1;
    for sheet in book.get_sheet_vec() {
        let _ = write_sheet::write(sheet, &dir, &shared_strings, index, &format_map)?;
        index = index + 1;
    }
    write_to_file(path, &dir)?;
    dir.close()?;
    Ok(())
}
