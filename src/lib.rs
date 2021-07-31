use std::env;
use calamine::{open_workbook, Xlsx, Reader, DataType};
use oracle::{Connection, Result};

pub fn read_write()->Vec<Result<String>> {
    let mut result = Vec::new();
    let args: Vec<String> = env::args().collect();

    let file = &args[1];
    println!("In file {}", file);

    let mut excel: Xlsx<_> = open_workbook(file).unwrap();
    //get different sheet names
    let sheets = excel.sheet_names().to_owned();
    let mut sheets_data = Vec::new();
    for s in &sheets {
        let mut data_by_row = Vec::new();
        if let Some(Ok(r)) = excel.worksheet_range(s){
            for row in r.rows(){
                let mut row_holder= Vec::new();
                for r_item in row {
                    match r_item {
                        DataType::String(x) => {
                            let x = x.replace(".", "_");
                            let x = x.replace("-", "_");
                            let x = x.replace(" ", "_");
                            row_holder.push(x.to_string())},
                        _ => row_holder.push("".to_string()),
    
                    }
                }
                data_by_row.push(row_holder);
            }
            //println!("{:?}",data_by_row);
        }
        sheets_data.push(data_by_row);

    }
    for (i,s) in sheets.iter().enumerate(){
        result.push(write_to_oracle(s, &mut sheets_data[i]));
    }
    result
}

pub fn write_to_oracle(sheet_name: &str, data_by_row: &mut Vec<Vec<String>>)->Result<String>{
    let conn = Connection::connect("test", "test", "localhost/xe")?;

    //create the table
    let mut sql = format!("create table {} ( ",sheet_name);
    for (i,data) in data_by_row[0].clone().iter().enumerate(){
        sql.push_str(&data);
        if i+1 >= data_by_row[0].len(){
            sql.push_str(" varchar2(1000) ");
        }
        else{
            sql.push_str(" varchar2(1000), ");
        }
    }
    sql.push_str(" ) ");

    println!("{}",sql);
    conn.execute(&sql, &[])?;
    conn.commit()?;

    let mut sql2 = format!("insert into {} values ( ",sheet_name);
    let row_0 = data_by_row.remove(0);
    for (i,_) in row_0.iter().enumerate(){
        if i+1 >= row_0.len(){
            let f = format!(":{} )",i+1);
            sql2.push_str(&f);
        }
        else{
            let f = format!(":{}, ",i+1);
            sql2.push_str(&f);
        }

    }
    println!("insert statement: {}",sql2);
    let mut stmt = conn.prepare(&sql2,&[])?;
    for row in data_by_row {
        for (i,ri) in row.iter().enumerate() {
            stmt.bind(i+1, ri)?;
            
        }
        stmt.execute(&[])?;
        conn.commit()?;
    }

    Ok("Success!".to_string())
}