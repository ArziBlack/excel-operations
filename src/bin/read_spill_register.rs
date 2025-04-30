use std::path::Path;
use std::error::Error;
use calamine::{open_workbook, Reader, Xlsx};

fn main() -> Result<(), Box<dyn Error>> {
    let path = Path::new("spill_register.xlsx");
    
    if !path.exists() {
        println!("spill_register.xlsx file not found!");
        return Ok(());
    }
    
    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    
    // Print all sheet names
    println!("Available sheets:");
    for sheet_name in workbook.sheet_names() {
        println!("  - {}", sheet_name);
    }
    
    // Try to read the first sheet regardless of name
    if let Some(sheet_name) = workbook.sheet_names().get(0).cloned() {
        println!("\nReading sheet: {}", sheet_name);
        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            // Find the row with actual headers (usually after any title rows)
            let mut header_row_index = 0;
            for (row_idx, row) in range.rows().enumerate() {
                // Print each row for examination
                println!("Row {}: {:?}", row_idx + 1, row);
                
                // If we find a row with multiple non-empty cells, it's likely the header row
                if row.iter().filter(|cell| !matches!(cell, calamine::DataType::Empty)).count() > 3 {
                    header_row_index = row_idx;
                    if row_idx > 0 { // Skip the first row as it might be a title
                        break;
                    }
                }
            }
            
            // Print the identified header row
            if let Some(headers) = range.rows().nth(header_row_index) {
                println!("\nIdentified Header Row (Row {}): {:?}", header_row_index + 1, headers);
                for (idx, header) in headers.iter().enumerate() {
                    if !matches!(header, calamine::DataType::Empty) {
                        println!("  Column {}: {:?}", idx + 1, header);
                    }
                }
            }
        }
    } else {
        println!("No sheets found in the workbook!");
    }
    
    Ok(())
}
