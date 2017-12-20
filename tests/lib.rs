// cargo test --all-features -- --nocapture

extern crate spsheet;
use spsheet::{Book,Sheet,Cell,column_to_index,index_to_column,column_and_row_to_index};
use spsheet::style::Style;

use std::path::Path;

#[cfg(feature = "ods")]
use spsheet::ods;

#[cfg(feature = "xlsx")]
use spsheet::xlsx;

fn make_sheet1() -> Sheet {
    let mut sheet = Sheet::new("シート1");
    sheet.add_cell(Cell::str("a"), 0, 0);
    sheet.add_cell(Cell::str("b"), 0, 1);
    sheet.add_cell(Cell::float(1.0), 1, 0);
    sheet.add_cell(Cell::float(2.0), 1, 1);
    sheet.add_cell(Cell::date_with_style("2017-12-02", Style::new("MM\\月DD\\日")), 2, 0);
    sheet.add_cell(Cell::date_with_style("2017-12-02T13:30:00", Style::new("YYYY/MM/DD\\ HH:MM:SS")), 2, 1);
    sheet
}

fn make_sheet2() -> Sheet {
    let mut sheet = Sheet::new("シート2");
    sheet.add_cell(Cell::str("予定表～①ﾊﾝｶｸだ"), 0, 0);
    sheet
}

fn make_sheet3() -> Sheet {
    let mut sheet = Sheet::new("シート3");
    sheet.add_cell(Cell::str("a"), 0, 0);
    sheet.add_cell(Cell::str("b"), 0, 2);
    sheet.add_cell(Cell::str("c"), 2, 0);
    sheet.add_cell(Cell::str("d"), 2, 2);
    sheet.add_cell(Cell::str("e"), 2, 4);
    sheet.add_cell(Cell::str("f"), 4, 0);
    sheet.add_cell(Cell::str("g"), 4, 4);
    sheet
}

fn make_sheet4() -> Sheet {
    Sheet::new("シート4")
}

fn make_book() -> Book {
    let mut book = Book::new();
    book.add_sheet(make_sheet1());
    book.add_sheet(make_sheet2());
    book.add_sheet(make_sheet3());
    book.add_sheet(make_sheet4());
    book
}

#[test]
fn it_works() {
    for i in vec![0,1,26,27,28,100,101,102] {
        assert_eq!(i, column_to_index(index_to_column(i)));
    }
    for i in vec!["A", "B", "Z", "AA", "AB", "ZZ", "AAA", "AAB", "ABC"] {
        assert_eq!(i, index_to_column(column_to_index(i)));
    }
    assert_eq!(Some((701,11)), column_and_row_to_index("ZZ12"));
}

#[test]
#[cfg(feature = "ods")]
fn ods_test() {
    let book = make_book();
    let _ = ods::write(&book, Path::new("./tests/test.ods"));
    let res = ods::read(Path::new("./tests/test.ods")).unwrap();
    assert_eq!(book, res);
}

#[test]
#[cfg(feature = "xlsx")]
fn xlsx_test() {
    let book = make_book();
    let _ = xlsx::write(&book, Path::new("./tests/test.xlsx"));
    let res = xlsx::read(Path::new("./tests/test.xlsx")).unwrap();
    assert_eq!(book, res);
}