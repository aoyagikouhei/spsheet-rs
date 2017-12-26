//! A xlsx or ods read and write library
extern crate chrono;
extern crate era_jp;

#[macro_use]
extern crate nom;

use chrono::prelude::*;
use std::collections::HashMap;
use std::borrow::Cow;

pub mod style;
use style::Style;

#[cfg(feature = "ods")]
pub mod ods;

#[cfg(feature = "xlsx")]
pub mod xlsx;

#[cfg(any(feature = "ods", feature = "xlsx"))]
mod file_common;

/// String index to usize index start with 0
///
/// ```
/// use spsheet::*;
/// assert_eq!(0, column_to_index("A"));
/// assert_eq!(1, column_to_index("B"));
/// assert_eq!(25, column_to_index("Z"));
/// assert_eq!(26, column_to_index("AA"));
/// ```
pub fn column_to_index<'a, S>(value: S) -> usize 
    where S: Into<Cow<'a, str>>
{
    let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let mut index = 0;
    for (i, c) in value.into().chars().enumerate() {
        if i != 0 {
            index = (index + 1) * 26;
        }
        index = index + alphabet.find(c).unwrap();
    }
    index
}

/// Usize index to String index
///
/// ```
/// use spsheet::*;
/// assert_eq!("A", index_to_column(0));
/// assert_eq!("B", index_to_column(1));
/// assert_eq!("Z", index_to_column(25));
/// assert_eq!("AA", index_to_column(26));
/// ```
pub fn index_to_column(index: usize) -> String {
    let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let mut result = String::from("");
    let mut work = index;
    loop {
        result.push(alphabet.chars().nth(work % 26).unwrap());
        if work < 26 {
            break;
        }
        work = (work / 26) - 1;
    }
    result.chars().rev().collect()
}

/// Column and row String index to usize index pair
///
/// ```
/// use spsheet::*;
/// assert_eq!(Some((701,11)), column_and_row_to_index("ZZ12"));
/// ```
pub fn column_and_row_to_index<'a, S>(value: S) -> Option<(usize, usize)>
    where S: Into<Cow<'a, str>>
{
    let val = String::from(value.into());
    let digit = "0123456789";
    let mut index = std::usize::MAX;
    for i in digit.chars() {
        match val.find(i) {
            Some(n) => {
                if index > n {
                    index = n;
                }
            },
            None => {}
        }
    }
    if std::usize::MAX == index {
        None
    } else {
        unsafe {
            Some(
                (
                    column_to_index(val.slice_unchecked(0, index)), 
                    val.slice_unchecked(index, val.len()).parse::<usize>().unwrap() - 1
                )
            )
        }
    }
}

/// Book has owner of sheets.
///
/// ```
/// let _ = spsheet::Book::new();
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Book {
    sheets: Vec<Sheet>,
}

impl Book {
    pub fn new() -> Book {
        Book {
            sheets: Vec::new()
        }
    }

    pub fn add_sheet(&mut self, sheet: Sheet) {
        self.sheets.push(sheet);
    }

    pub fn get_sheet(&self, index: usize) -> &Sheet {
        &self.sheets[index]
    }

    pub fn get_sheet_size(&self) -> usize {
        self.sheets.len() as usize
    }

    pub fn get_sheet_vec(&self) -> &Vec<Sheet> {
        &self.sheets
    }
}

/// Sheet has owner of cells.
///
/// ```
/// let _ = spsheet::Sheet::new("sheet1");
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Sheet {
    name: String,
    rows: HashMap<usize, HashMap<usize, Cell>>,
}

impl Sheet {
    pub fn new<'a, S>(name: S) -> Sheet 
        where S: Into<Cow<'a, str>>
    {
        Sheet {
            name: name.into().into_owned(),
            rows: HashMap::new()
        }
    }

    pub fn set_name<'a, S>(&mut self, name: S)
        where S: Into<Cow<'a, str>>
    {
        self.name = name.into().into_owned();
    }

    pub fn get_name(&self) -> &String {
        &self.name
    }

    pub fn add_cell(&mut self, cell: Cell, row_index: usize, column_index: usize) {
        if let Some(row) = self.rows.get_mut(&row_index) {
            row.insert(column_index, cell);
            return;
        }        
        let mut row = HashMap::new();
        row.insert(column_index, cell);
        self.rows.insert(row_index, row);
    }

    pub fn get_cell(&self, row_index: usize, column_index: usize) -> Option<&Cell> {
        if let Some(row) = self.rows.get(&row_index) {
            return row.get(&column_index);
        } else {
            return None;
        }
    }

    pub fn get_rows(&self) -> &HashMap<usize, HashMap<usize, Cell>> {
        &self.rows
    }

    pub fn sorted_access<F>(&self, mut callback: F) 
        where F : FnMut(usize, usize, &Cell) -> () 
    {
        let mut row_index_vec: Vec<usize> = Vec::new();
        for (row_index, _) in self.get_rows().iter() {
            row_index_vec.push(*row_index);
        }
        row_index_vec.sort();
        for row_index in row_index_vec {
            let mut column_index_vec: Vec<usize> = Vec::new();
            let columns = self.get_rows().get(&row_index).unwrap();
            for (column_index, _) in columns.iter() {
                column_index_vec.push(*column_index);
            }
            column_index_vec.sort();
            for column_index in column_index_vec {
                let cell = columns.get(&column_index).unwrap();
                callback(row_index, column_index, cell);
            }
        }
    }

    pub fn walk_through<F>(&self, mut callback: F) 
        where F : FnMut(usize, usize, &Cell) -> () 
    {
        for (&row_index, rows) in self.get_rows() {
            for (&col_index, cell) in rows {
                callback(row_index, col_index, cell);
            }
        }
    }

    pub fn get_max_index(&self) -> Option<(usize, usize)> {
        if self.get_rows().len() == 0 {
            return None;
        }
        let mut max_row_index = 0;
        let mut max_column_index = 0;
        for (row_index, columns) in self.get_rows().iter() {
            if max_row_index < *row_index {
                max_row_index = *row_index;
            } 
            for (column_index, _) in columns.iter() {
                if max_column_index < *column_index {
                    max_column_index = *column_index;
                } 
            }
        }
        Some((max_row_index, max_column_index))
    }
}

/// Cell has owner of value and style.
///
/// ```
/// let _ = spsheet::Cell::str("value");
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    value: Value,
    style: Style,
}

impl Cell {
    pub fn new(value: Value, style: Style) -> Cell
    {
        Cell {
            value,
            style,
        }
    }

    pub fn str<'a, S>(value: S) -> Cell 
        where S: Into<Cow<'a, str>>
    {
        Cell::new(
            Value::Str(value.into().into_owned()),
            Style::new(""))
    }

    pub fn float(value: f64) -> Cell 
    {
        Cell::new(
            Value::Float(value),
            Style::new(""))
    }
/*
    pub fn date<'a, S>(value: S) -> Cell 
        where S: Into<Cow<'a, str>>
    {
        Cell::date_with_style(value, Style::new("%Y-%m-%d"))
    }
*/
    pub fn date_with_style<'a, S>(value: S, style: Style) -> Cell 
        where S: Into<Cow<'a, str>>
    {
        let mut str_value = value.into().into_owned();
        let postfix = match str_value.find('T') {
            Some(_) => "Z",
            None => "T00:00:00Z"
        };
        str_value.push_str(postfix);
        Cell::new(
            Value::Date(str_value.parse::<DateTime<Utc>>().unwrap()),
            style
        )
    }

    pub fn get_value(&self) -> &Value {
        &self.value
    }

    pub fn get_style(&self) -> &Style {
        &self.style
    }

    pub fn set_style(&mut self, style: Style) {
        self.style = style;
    }

    pub fn get_formated_value(&self) -> Option<String> {
        match self.value {
            Value::Date(dt) => {
                self.style.get_formated_date(&dt)
            },
            _ => None,
        }
    }
}

/// Value has Str, Float, Data value.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// String Value
    Str(String),
    /// Float Value
    Float(f64),
    /// Data Value
    Date(DateTime<Utc>),
    /// Currency Value
    Currency(f64),
}