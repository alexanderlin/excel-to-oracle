### excel-to-oracle

converts an excel sheet (tested with xlsx) to an oracle table (with the data type of every column as a string) with the sheet name as the table name. 

# dependencies:

calamine: https://docs.rs/crate/calamine/0.18.0

oracle: https://docs.rs/crate/oracle/0.5.2

# How to run: 
```
cargo run path column_capacity username password database_connection
```
path: path to the excel file to convert with the file included
column_capacity: how much space to allocate to varchar2() when the table is being created
username: username to database
password: password to database
database_connection: connection to database
