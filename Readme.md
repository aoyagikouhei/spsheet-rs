# spsheet-rs

A xlsx or ods read and write library

[Documentation](https://docs.rs/spsheet-rs)

## Description

**spsheet-rs** is is a pure Rust library to read and write any xlsx and ods file.

## Features
- [x] xlsx Read
- [x] xlsx Write
- [x] ods Read
- [x] ods Write
- [x] Cell Value
- [-] Cell Format(Date)
- [ ] Cell Format(Digit)
- [ ] Cell Border
- [ ] Cell Color
- [ ] Cell Width
- [ ] Cell Hegiht
- [ ] Formular

## Examples

```rust
extern crate spsheet;
use spsheet::ods;
use spsheet::xlsx;
use spsheet::{Book,Sheet,Cell};
use spsheet::style::Style;
use std::path::Path;

let mut book = Book::new();
let mut sheet = Sheet::new("シート1");
sheet.add_cell(Cell::str("a"), 0, 0);
sheet.add_cell(Cell::str("b"), 0, 1);
sheet.add_cell(Cell::float(1.0), 1, 0);
sheet.add_cell(Cell::float(2.0), 1, 1);
sheet.add_cell(Cell::date_with_style("2017-12-02", Style::new("MM\\月DD\\日")), 2, 0);
sheet.add_cell(Cell::date_with_style("2017-12-02T13:30:00", Style::new("YYYY/MM/DD\\ HH:MM:SS")), 2, 1);
book.add_sheet(sheet);

let _ = ods::write(&book, Path::new("./tests/test.ods"));
let res = ods::read(Path::new("./tests/test.ods")).unwrap();
assert_eq!(book, res);

let _ = xlsx::write(&book, Path::new("./tests/test.xlsx"));
let res = xlsx::read(Path::new("./tests/test.xlsx")).unwrap();
assert_eq!(book, res);
```

If I can find Google Shreadsheet API Library in Rust, I'll support Google Shpeadsheet. 