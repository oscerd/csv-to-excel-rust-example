use structopt::StructOpt;

/// A basic example
#[derive(StructOpt, Debug)]
#[structopt(name = "basic")]
pub struct Opt {

    /// CSV Input file name
    #[structopt()]
    input: String,

    /// XLSX Output file name
    #[structopt()]
    output: String,

    // Sheet name
    /// Sheet Name
    #[structopt(default_value = "Worksheet")]
    sheet_name: String,
}


pub fn convert_csv_to_excel(
    opt: &Opt
) -> Result<(), Box<dyn std::error::Error>> {
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_path(&*opt.input)?;

    let mut wb = simple_excel_writer::Workbook::create(&*opt.output);
    let mut sheet = wb.create_sheet(&*opt.sheet_name);
    wb.write_sheet(&mut sheet, |sw| {
        for result in rdr.records() {
            if let Ok(record) = result {
                let mut row = simple_excel_writer::Row::new();
                for field in record.iter() {
                    row.add_cell(field);
                }
                sw.append_row(row)?;
            }
        }
        Ok(())
    })?;

    wb.close()?;
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let opt = Opt::from_args();
    println!("{:?}", opt);
    convert_csv_to_excel(&opt)
}
