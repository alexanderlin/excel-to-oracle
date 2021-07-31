use std::env;
use calamine::{open_workbook, Xlsx, Reader, DataType};
use oracle::{Connection, Result};

pub fn read_file()->Result<String> {
    let mut data_by_row = Vec::new();
    let args: Vec<String> = env::args().collect();

    let file = &args[1];
    println!("In file {}", file);

    let mut excel: Xlsx<_> = open_workbook(file).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1"){
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
    connect_to_oracle(&mut data_by_row)
}

pub fn connect_to_oracle(data_by_row: &mut Vec<Vec<String>>)->Result<String>{
    let conn = Connection::connect("test", "test", "localhost/xe")?;

    //create the table
    let mut sql = "create table Table1 ( ".to_owned();
    for (i,data) in data_by_row[0].clone().iter().enumerate(){
        sql.push_str(&data);
        if i+1 >= data_by_row[0].len(){
            sql.push_str(" varchar2(200) ");
        }
        else{
            sql.push_str(" varchar2(200), ");
        }
    }
    sql.push_str(" ) ");

    println!("{}",sql);
    conn.execute(&sql, &[])?;
    conn.commit()?;

    let mut sql2 = "insert into Table1 values ( ".to_owned();
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
            println!("{}",i);
            stmt.bind(i+1, ri)?;
            
        }
        stmt.execute(&[])?;
        conn.commit()?;
    }

    Ok("Success!".to_string())
}