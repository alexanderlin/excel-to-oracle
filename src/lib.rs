use std::env;
use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};

pub fn read_file(){
    let args: Vec<String> = env::args().collect();

    let file = &args[1];
    println!("In file {}", file);

    let mut excel: Xlsx<_> = open_workbook(file).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1"){
        // for row in r.rows(){
        //     println!("row={:?}, row[0]={:?}", row, row[0]);
        // }
        println!("row={:?}, row[0]={:?}", r.row, r.row[0]);

    }
}