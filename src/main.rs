use std::process;

pub fn convert_csv_to_excel(
    csv_path: &str,
    excel_path: &str,
    sheet_name: &str
) -> Result<(), Box<dyn std::error::Error>> {
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_path(csv_path)?;

    let mut wb = simple_excel_writer::Workbook::create(excel_path);
    let mut sheet = wb.create_sheet(&mut &sheet_name);

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
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        println!("A CSV file as input and an output file must be provided");
        process::exit(1)
    } ;
    let mut sheet_name = "Worksheet";
    if args.len() >= 4 {
        sheet_name = &args[3];
    }
    convert_csv_to_excel(&args[1], &args[2], &mut sheet_name)
}
