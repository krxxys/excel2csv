use calamine::{open_workbook_auto, DataType, Range, Reader, XlsOptions};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;


fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");
    let output_dir = env::args()
        .nth(2)
        .expect("Expecting a sheet name as second argument");

    let pb_from_file = PathBuf::from(file);
    match pb_from_file.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

   
    let mut xl = open_workbook_auto(&pb_from_file).unwrap();
    let fileoutput_sce = PathBuf::from(output_dir);
    let sheets = xl.sheet_names().to_owned();
    for s in sheets {

        let mut new_s = String::new(); 
        new_s.push_str(&fileoutput_sce.to_str().unwrap());
        new_s.push_str(&s.clone());
        new_s.push_str(".csv");

        let mut dest = BufWriter::new(File::create(&new_s).unwrap());
        let range = xl.worksheet_range(&s).unwrap().unwrap();
        write_range(&mut dest, &range).unwrap();
    }
    
}

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "{}", s),
                DataType::Float(ref f) => {write!(dest, "{}", f)},
                DataType::DateTime(ref f) => write!(dest, "{}", DataType::as_date(c).unwrap().format("%d/%m/%y")),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ",")?;
            }
        }
        write!(dest, "\r\n")?;
 
    }
    Ok(())
}