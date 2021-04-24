use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;
use std::time::Instant;

fn main() {
    let start = Instant::now();
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = sce.with_extension("csv");
    let output_name = sce.as_os_str();
    let name_str = output_name.to_str().unwrap().split('.').collect::<Vec<_>>()[0];
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&sce).unwrap();
    let sheet = xl.sheet_names()[0].to_owned();
    let range = xl.worksheet_range(sheet.as_str()).unwrap().unwrap();

    write_range(&mut dest, &range).unwrap();

    let duration = start.elapsed();

    println!("Time elapsed in xlsx_to_csv is: {:?}", duration);
    println!("Output file was saved in {}.csv", name_str);
}

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let sep = env::args().nth(2).expect("Please provide an seperator");
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "{}", s.replace("\n", "")),
                DataType::Float(ref f) | DataType::DateTime(ref f) => write!(dest, "{}", f),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, "{}", sep)?;
            }
        }
        write!(dest, "\n")?;
    }
    Ok(())
}
