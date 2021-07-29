use std::env;
use calamine::{open_workbook, Xlsx, Reader};
use oracle::{Connection, Result};

pub fn read_file(){
    let args: Vec<String> = env::args().collect();

    let file = &args[1];
    println!("In file {}", file);

    let mut excel: Xlsx<_> = open_workbook(file).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1"){
        for row in r.rows(){
            //println!("row={:?}, row[0]={:?}", row, row[0]);
            for r_item in row {
                println!("row item: {:?}",r_item);
            }
        }
        

    }
}

pub fn connect_to_oracle()->Result<String>{
    let conn = Connection::connect("test", "test", "localhost/xe")?;
    conn.execute("create table person (id number(38), name varchar2(40))", &[])?;
    conn.commit()?;
    Ok("Success!".to_string())
}