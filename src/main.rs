use excel_to_oracle::{read_file, connect_to_oracle};

fn main() {
    read_file();
    println!("now connecting to oracle database");
    println!("{:?}",connect_to_oracle());
}