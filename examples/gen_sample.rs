use anyhow::Result;

fn main() -> Result<()> {
    let mut book = umya_spreadsheet::new_file();

    let sheet = book.get_active_sheet_mut();

    // Make row1 non-empty so the main program's maxColumn logic matches Go behavior.
    sheet.get_cell_mut("A1").set_value("header");

    // Put a value in row3 with parentheses to trigger replacement + red fill.
    sheet.get_cell_mut("A3").set_value("hello(world)");

    umya_spreadsheet::writer::xlsx::write(&book, "sample.xlsx")?;
    println!("Wrote sample.xlsx");
    Ok(())
}
