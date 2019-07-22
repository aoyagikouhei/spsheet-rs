use file_common::*;
use std::io::Cursor;
use std::result;
use super::chrono::prelude::*;
use super::quick_xml::events::{Event, BytesDecl};
use super::quick_xml::Writer;
use super::tempdir::TempDir;
use super::{Sheet, Value, index_to_column};
use super::XlsxError;
use std::collections::HashMap;

pub fn write(sheet: &Sheet, dir: &TempDir, shared_strings: &HashMap<String, usize>, index: usize, format_map: &HashMap<String, usize>) -> result::Result<(), XlsxError> {
    let dimension = match sheet.get_max_index() {
        Some((max_row_index, max_column_index)) => {
            if max_row_index == 0 && max_column_index == 0 {
                String::from("A1")
            } else {
                format!("A1:{}{}", index_to_column(max_column_index), (max_row_index + 1).to_string())
            }
        },
        None => String::from("A1")
    };
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let _ = writer.write_event(Event::Decl(
        BytesDecl::new(b"1.0", Some(b"UTF-8"), Some(b"yes"))));
    write_text_node(&mut writer, "\n");
    write_start_tag(&mut writer, "worksheet", vec![
        ("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
        ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ], false);
    write_start_tag(&mut writer, "sheetPr", vec![("filterMode", "false")], false);
    write_start_tag(&mut writer, "pageSetUpPr", vec![("fitToPage", "false")], true);
    write_end_tag(&mut writer, "sheetPr");
    write_start_tag(&mut writer, "dimension", vec![("ref", dimension.as_str())], true);
    write_start_tag(&mut writer, "sheetViews", vec![], false);
    write_start_tag(&mut writer, "sheetView", vec![("showFormulas", "false"),("showGridLines", "true"),("showRowColHeaders", "true"),("showZeros", "true"),("rightToLeft", "false"),("tabSelected", "true"),("showOutlineSymbols", "true"),("defaultGridColor", "true"),("view", "normal"),("topLeftCell", "A1"),("colorId", "64"),("zoomScale", "100"),("zoomScaleNormal", "100"),("zoomScalePageLayoutView", "100"),("workbookViewId", "0")], false);
    write_start_tag(&mut writer, "selection", vec![("pane", "topLeft"),("activeCell", "A1"),("activeCellId", "0"),("sqref", "A1")], true);
    write_end_tag(&mut writer, "sheetView");
    write_end_tag(&mut writer, "sheetViews");
    write_start_tag(&mut writer, "sheetFormatPr", vec![("defaultRowHeight", "12.8"),("zeroHeight", "false"),("outlineLevelRow", "0"),("outlineLevelCol", "0")], true);
    write_start_tag(&mut writer, "cols", vec![], false);
    write_start_tag(&mut writer, "col", vec![("collapsed", "false"),("customWidth", "true"),("hidden", "false"),("outlineLevel", "0"),("max", "1025"),("min", "1"),("style", "0"),("width", "10.86")], true);
    write_end_tag(&mut writer, "cols");
    if sheet.get_rows().len() == 0 {
        write_start_tag(&mut writer, "sheetData", vec![], true);
    } else {
        write_start_tag(&mut writer, "sheetData", vec![], false);
        let mut current_row_index = ::std::usize::MAX;
        sheet.sorted_access(|row_index, column_index, cell| {
            let row_str = (row_index + 1).to_string();
            if current_row_index != row_index {
                if current_row_index != ::std::usize::MAX {
                    write_end_tag(&mut writer, "row");
                }
                current_row_index = row_index;
                write_start_tag(&mut writer, "row", vec![
                    ("r", &row_str),
                    ("customFormat", "false"),
                    ("hidden", "false"),
                    ("customHeight", "false"),
                    ("outlineLevel", "0"),
                    ("collapsed", "false"),
                ], false);
            }
            let col_str = format!(
                "{}{}", index_to_column(column_index), row_str);
            let format = cell.get_format().get_content();
            match cell.get_value() {
                &Value::Str(ref val) => {
                    write_start_tag(&mut writer, "c", vec![
                        ("r", &col_str),
                        ("s", "0"),
                        ("t", "s"),
                    ], false);
                    write_start_tag(&mut writer, "v", vec![], false);
                    let val_index = shared_strings.get(val).unwrap().to_string();
                    write_text_node(&mut writer, val_index.as_str());
                },
                &Value::Float(ref val) => {
                    write_start_tag(&mut writer, "c", vec![
                        ("r", &col_str),
                        ("s", "0"),
                        ("t", "n"),
                    ], false);
                    write_start_tag(&mut writer, "v", vec![], false);
                    write_text_node(&mut writer, val.to_string().as_str());
                },
                &Value::Date(ref val) => {
                    let s_value = format_map.get(format).unwrap();
                    write_start_tag(&mut writer, "c", vec![
                        ("r", &col_str),
                        ("s", s_value.to_string().as_str()),
                        ("t", "n"),
                    ], false);
                    write_start_tag(&mut writer, "v", vec![], false);
                    write_text_node(&mut writer, datetime_to_serail(val).to_string().as_str());
                },
                &Value::Currency(ref val) => {
                    let s_value = format_map.get(format).unwrap();
                    write_start_tag(&mut writer, "c", vec![
                        ("r", &col_str),
                        ("s", s_value.to_string().as_str()),
                        ("t", "n"),
                    ], false);
                    write_start_tag(&mut writer, "v", vec![], false);
                    write_text_node(&mut writer, val.to_string().as_str());
                },
            }
            write_end_tag(&mut writer, "v");
            write_end_tag(&mut writer, "c");
        });
        write_end_tag(&mut writer, "row");
        write_end_tag(&mut writer, "sheetData");
    }
    write_start_tag(&mut writer, "printOptions", vec![("headings", "false"),("gridLines", "false"),("gridLinesSet", "true"),("horizontalCentered", "false"),("verticalCentered", "false")], true);
    write_start_tag(&mut writer, "pageMargins", vec![("left", "0.7875"),("right", "0.7875"),("top", "1.025"),("bottom", "1.025"),("header", "0.7875"),("footer", "0.7875")], true);
    write_start_tag(&mut writer, "pageSetup", vec![("paperSize", "9"),("scale", "100"),("firstPageNumber", "1"),("fitToWidth", "1"),("fitToHeight", "1"),("pageOrder", "downThenOver"),("orientation", "portrait"),("blackAndWhite", "false"),("draft", "false"),("cellComments", "none"),("useFirstPageNumber", "true"),("horizontalDpi", "300"),("verticalDpi", "300"),("copies", "1")], true);
    write_start_tag(&mut writer, "headerFooter", vec![("differentFirst", "false"),("differentOddEven", "false")], false);
    write_start_tag(&mut writer, "oddHeader", vec![], false);
    write_text_node(&mut writer, "&amp;C&amp;&quot;Arial,標準&quot;&amp;A");
    write_end_tag(&mut writer, "oddHeader");
    write_start_tag(&mut writer, "oddFooter", vec![], false);
    write_text_node(&mut writer, "&amp;C&amp;&quot;Arial,標準&quot;ページ &amp;P");
    write_end_tag(&mut writer, "oddFooter");
    write_end_tag(&mut writer, "headerFooter");
    write_end_tag(&mut writer, "worksheet");
    let _ = make_file_from_writer(format!("xl/worksheets/sheet{}.xml", index).as_str(), dir, writer, Some("xl/worksheets"))?;
    Ok(())
}

fn datetime_to_serail(src: &DateTime<Utc>) -> f64 {
    let seconds = src.hour() * 3600 + src.minute() * 60 + src.second();
    let base = Utc.ymd(1900, 1, 1).and_hms(0, 0, 0);
    let days = src.num_days_from_ce() - base.num_days_from_ce() + 2;
    days as f64 + seconds as f64 / 86400.0
}
