use std::path::Path;
use std::error::Error;
use chrono::{Local, DateTime};
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Format, Color, FormatBorder};

/// Represents a spill incident to be recorded in the register
pub struct SpillIncident {
    pub date_time: DateTime<Local>,
    pub location: String,
    pub material: String,
    pub quantity_liters: f64,
    pub reported_by: String,
    pub cleanup_status: String,
    pub notes: Option<String>,
}

/// Adds a new spill incident to the spill register
pub fn add_spill_incident(incident: SpillIncident) -> Result<(), Box<dyn Error>> {
    let file_path = Path::new("spill_register.xlsx");
    
    // Check if file exists, if not create it with headers
    // if !file_path.exists() {
    //     create_spill_register()?;
    // }
    
    // First read the existing file to find the next empty row
    let next_row = find_next_empty_row(file_path)?;
    
    // Now open the file for writing
    let mut workbook = open_existing_workbook(file_path)?;
    
    // Format the date as string (since we can't use DateTime directly)
    let date_str = incident.date_time.format("%Y-%m-%d %H:%M:%S").to_string();
    
    // Get a mutable reference to the first worksheet
    // We need to use this approach since rust_xlsxwriter's API requires a mutable worksheet
    let worksheet = &mut workbook.worksheets_mut()[0];

    // Write the data to the next empty row
    worksheet.write_string(next_row, 0, &date_str)?;
    worksheet.write_string(next_row, 1, &incident.location)?;
    worksheet.write_string(next_row, 2, &incident.material)?;
    worksheet.write_number(next_row, 3, incident.quantity_liters)?;
    worksheet.write_string(next_row, 4, &incident.reported_by)?;
    worksheet.write_string(next_row, 5, &incident.cleanup_status)?;
    
    // Write notes if provided
    if let Some(notes) = &incident.notes {
        worksheet.write_string(next_row, 6, notes)?;
    }
    
    // Save the workbook
    workbook.save("spill_register.xlsx")?;
    println!("Spill incident added successfully at row {}", next_row + 1);
    
    Ok(())
}

/// Finds the next empty row in the spill register
fn find_next_empty_row(file_path: &Path) -> Result<u32, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;
    
    if let Ok(range) = workbook.worksheet_range("Spill Register") {
        // Get the number of rows in the sheet (add 1 for the next empty row)
        let row_count = range.height() as u32;
        
        // Return the next row index (0-based)
        Ok(row_count)
    } else {
        // If sheet doesn't exist or is empty, start at row 1 (after header)
        Ok(1)
    }
}

/// Opens an existing workbook for writing
fn open_existing_workbook(file_path: &Path) -> Result<Workbook, Box<dyn Error>> {
    // We need to create a new workbook and copy the existing data
    // since rust_xlsxwriter doesn't support direct modification
    let mut original_workbook: Xlsx<_> = open_workbook(file_path)?;
    let mut new_workbook = Workbook::new();
    
    // Get the data from the original workbook
    if let Ok(range) = original_workbook.worksheet_range("Spill Register") {
        let worksheet = new_workbook.add_worksheet().set_name("Spill Register")?;
        
        // Copy all existing data
        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                match cell {
                    calamine::DataType::String(s) => {
                        worksheet.write_string(row_idx as u32, col_idx as u16, s)?;
                    },
                    calamine::DataType::Float(f) => {
                        worksheet.write_number(row_idx as u32, col_idx as u16, *f)?;
                    },
                    calamine::DataType::Int(i) => {
                        worksheet.write_number(row_idx as u32, col_idx as u16, *i as f64)?;
                    },
                    calamine::DataType::DateTime(d) => {
                        // Convert calamine DateTime to string
                        worksheet.write_number(row_idx as u32, col_idx as u16, *d)?;
                    },
                    calamine::DataType::Bool(b) => {
                        worksheet.write_boolean(row_idx as u32, col_idx as u16, *b)?;
                    },
                    _ => {
                        // For empty cells or other types, write an empty string
                        worksheet.write_string(row_idx as u32, col_idx as u16, "")?;
                    }
                }
            }
        }
        
        // Set column widths for better readability
        worksheet.set_column_width(0, 20.0)?; // DateTime
        worksheet.set_column_width(1, 20.0)?; // Location
        worksheet.set_column_width(2, 15.0)?; // Material
        worksheet.set_column_width(3, 15.0)?; // Quantity
        worksheet.set_column_width(4, 15.0)?; // Reported By
        worksheet.set_column_width(5, 15.0)?; // Cleanup Status
        worksheet.set_column_width(6, 30.0)?; // Notes
        
        Ok(new_workbook)
    } else {
        // If the sheet doesn't exist, create a new one
        create_spill_register()
    }
}

/// Creates a new spill register with headers
fn create_spill_register() -> Result<Workbook, Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet().set_name("Spill Register")?;
    
    // Create header format
    let header_format = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0x4F81BD))
        .set_font_color(Color::RGB(0xFFFFFF))
        .set_border(FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center);
    
    // Write headers
    worksheet.write_with_format(0, 0, "Date & Time", &header_format)?;
    worksheet.write_with_format(0, 1, "Location", &header_format)?;
    worksheet.write_with_format(0, 2, "Material", &header_format)?;
    worksheet.write_with_format(0, 3, "Quantity (L)", &header_format)?;
    worksheet.write_with_format(0, 4, "Reported By", &header_format)?;
    worksheet.write_with_format(0, 5, "Cleanup Status", &header_format)?;
    worksheet.write_with_format(0, 6, "Notes", &header_format)?;
    
    // Set column widths for better readability
    worksheet.set_column_width(0, 20.0)?; // DateTime
    worksheet.set_column_width(1, 20.0)?; // Location
    worksheet.set_column_width(2, 15.0)?; // Material
    worksheet.set_column_width(3, 15.0)?; // Quantity
    worksheet.set_column_width(4, 15.0)?; // Reported By
    worksheet.set_column_width(5, 15.0)?; // Cleanup Status
    worksheet.set_column_width(6, 30.0)?; // Notes
    
    // Save the workbook
    workbook.save("spill_register.xlsx")?;
    println!("Created new spill register file: spill_register.xlsx");
    
    Ok(workbook)
}
