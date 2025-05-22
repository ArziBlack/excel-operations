use std::path::Path;
use std::error::Error;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Format, Color, FormatBorder};

fn main() -> Result<(), Box<dyn Error>> {
    // Create a sample file first if it doesn't exist
    let path = Path::new("data.xlsx");
    if !path.exists() {
        println!("Sample file 'data.xlsx' not found. Creating it first...");
        create_sample_file()?;
    }

    // Example 1: Reading an Excel file
    println!("Example 1: Reading an Excel file");
    read_excel_example()?;

    // Example 2: Creating and writing to an Excel file
    println!("\nExample 2: Creating and writing to an Excel file");
    write_excel_example()?;

    // Example 3: Reading, modifying, and writing an Excel file
    println!("\nExample 3: Reading, modifying, and writing an Excel file");
    modify_excel_example()?;

    // Load and process the spill data
    load_and_process_spill_data()?;

    println!("\nAll examples completed successfully!");
    println!("Files created:");
    println!("- data.xlsx (sample data)");
    println!("- output.xlsx (from write example)");
    println!("- modified_data.xlsx (from modify example)");
    println!("- loaded_register.xlsx (from load and process spill data)");

    Ok(())
}

fn read_excel_example() -> Result<(), Box<dyn Error>> {
    // Path to the Excel file
    let path = Path::new("data.xlsx");
    
    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    
    // Select a worksheet by name
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        // Iterate through rows and cells
        for (row_index, row) in range.rows().enumerate() {
            println!("Row {}: {:?}", row_index + 1, row);
            
            // Access individual cells
            for (col_index, cell) in row.iter().enumerate() {
                println!("  Cell ({}, {}): {:?}", row_index + 1, col_index + 1, cell);
            }
        }
        
        // Access a specific cell by coordinates (0-based)
        if let Some(cell) = range.get_value((1, 1)) {
            println!("Value at B2: {:?}", cell);
        }
    } else {
        println!("Sheet 'Sheet1' not found or error occurred");
    }
    
    Ok(())
}

fn write_excel_example() -> Result<(), Box<dyn Error>> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    
    // Add a worksheet
    let worksheet = workbook.add_worksheet().set_name("Sheet1")?;
    
    // Create a header format
    let header_format = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xD8E4BC))
        .set_border(FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center);
    
    // Write headers with formatting
    worksheet.write_with_format(0, 0, "Name", &header_format)?;
    worksheet.write_with_format(0, 1, "Age", &header_format)?;
    worksheet.write_with_format(0, 2, "City", &header_format)?;
    
    // Write data
    worksheet.write_string(1, 0, "John Doe")?;
    worksheet.write_number(1, 1, 30.0)?;
    worksheet.write_string(1, 2, "New York")?;
    
    worksheet.write_string(2, 0, "Jane Smith")?;
    worksheet.write_number(2, 1, 25.0)?;
    worksheet.write_string(2, 2, "London")?;
    
    // Adjust column widths
    worksheet.set_column_width(0, 15.0)?;
    worksheet.set_column_width(1, 10.0)?;
    worksheet.set_column_width(2, 15.0)?;
    
    // Save the workbook
    workbook.save("output.xlsx")?;
    println!("Excel file created successfully: output.xlsx");
    
    Ok(())
}

fn modify_excel_example() -> Result<(), Box<dyn Error>> {
    // First, read the existing data
    let path = Path::new("data.xlsx");
    
    // Open the workbook for reading
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    
    // Read data from the worksheet
    let mut data = Vec::new();
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        for row in range.rows() {
            let row_data: Vec<String> = row.iter()
                .map(|cell| cell.to_string())
                .collect();
            data.push(row_data);
        }
    } else {
        println!("Sheet 'Sheet1' not found or error occurred");
        return Ok(());
    }
    
    // Modify the data (for example, add a new column with calculated values)
    for row in data.iter_mut().skip(1) { // Skip header row
        if row.len() >= 2 {
            // Try to parse the second column as a number and add 5 to it
            if let Ok(num) = row[1].parse::<f64>() {
                let new_value = num + 5.0;
                row.push(new_value.to_string());
            } else {
                row.push("N/A".to_string());
            }
        }
    }
    
    // Add header for the new column if there's data
    if !data.is_empty() {
        if data[0].len() >= 2 {
            data[0].push("Age + 5".to_string());
        }
    }
    
    // Create a new workbook for writing the modified data
    let mut new_workbook = Workbook::new();
    let worksheet = new_workbook.add_worksheet().set_name("Sheet1")?;
    
    // Write the modified data to the new workbook
    for (row_idx, row) in data.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            // Try to parse as number first
            if let Ok(num) = cell.parse::<f64>() {
                worksheet.write_number(row_idx as u32, col_idx as u16, num)?;
            } else {
                worksheet.write_string(row_idx as u32, col_idx as u16, cell)?;
            }
        }
    }
    
    // Save the modified workbook
    new_workbook.save("modified_data.xlsx")?;
    println!("Modified Excel file created successfully: modified_data.xlsx");
    
    Ok(())
}

fn create_sample_file() -> Result<(), Box<dyn Error>> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    
    // Add a worksheet
    let worksheet = workbook.add_worksheet().set_name("Sheet1")?;
    
    // Write headers
    worksheet.write_string(0, 0, "Name")?;
    worksheet.write_string(0, 1, "Age")?;
    worksheet.write_string(0, 2, "City")?;
    
    // Write data
    worksheet.write_string(1, 0, "John Doe")?;
    worksheet.write_number(1, 1, 30.0)?;
    worksheet.write_string(1, 2, "New York")?;
    
    worksheet.write_string(2, 0, "Jane Smith")?;
    worksheet.write_number(2, 1, 25.0)?;
    worksheet.write_string(2, 2, "London")?;
    
    // Save the workbook
    workbook.save("data.xlsx")?;
    println!("Sample Excel file created successfully: data.xlsx");
    
    Ok(())
}

fn load_and_process_spill_data() -> Result<(), Box<dyn Error>> {
    // Path to the Excel file
    let path = Path::new("Oil_Spill_2013.xlsx");
    
    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    println!("Workbook loaded successfully for Oil_Spill_2013.xlsx");

    // Print all sheet names for debugging
    let sheet_names = workbook.sheet_names();
    println!("Available sheets: {:?}", sheet_names);

    // Select a worksheet by name
    if let Ok(range) = workbook.worksheet_range("Oil_Spill_2013") {
        println!("Worksheet selected with {} rows", range.rows().count());
        // Prepare a new workbook for writing
        let mut new_workbook = Workbook::new();
        let worksheet = new_workbook.add_worksheet();

        // Write headers
        let headers = ["ID", "Spill Serial No", "Facility or Equipment", "Location", "Incident Date", "Coordinates", "Latitude", "Longitude", "Spill Status", "Description of Spill Causes", "Estimated Qty Spilled or Leaked", "Severity", "Extent of Pollution", "Precaution Measures", "Client ID", "OPL/OML No", "Date of Incident Observed", "Date of Incident Occurred", "Incident Cause", "Incident Type", "Nearest Town", "Operational Area", "Other Cause", "Other Type", "State", "Time of Incident Observed", "Time of Incident Occurred", "Type of Operation", "Coordinate Format", "Created At", "Updated At"];
        for (col, header) in headers.iter().enumerate() {
            worksheet.write_string(0, col as u16, *header)?;
        }

        // Iterate through rows and cells
        for (row_index, row) in range.rows().enumerate() {
            println!("Processing row {}: {:?}", row_index, row);
            if row_index == 0 {
                continue; // Skip header row
            }

            // Extract values matching SpillIncident struct
            let id = row.get(0).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let spill_serial_no = row.get(1).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let facility_or_equipment = row.get(2).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let location = row.get(3).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let incident_date = row.get(4).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let coordinates = row.get(5).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let latitude = row.get(6).and_then(|c| c.get_float()).unwrap_or(0.0);
            let longitude = row.get(7).and_then(|c| c.get_float()).unwrap_or(0.0);
            let spill_status = row.get(8).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let description_of_spill_causes = row.get(9).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let estimated_qty_spilled_or_leaked = row.get(10).and_then(|c| c.get_float()).unwrap_or(0.0);
            let severity = row.get(11).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let extent_of_pollution = row.get(12).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let precaution_measures = row.get(13).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let client_id = row.get(14).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let opl_oml_no = row.get(15).and_then(|c| c.get_int()).unwrap_or(0);
            let date_of_incident_observed = row.get(16).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let date_of_incident_occurred = row.get(17).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let incident_cause = row.get(18).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let incident_type = row.get(19).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let nearest_town = row.get(20).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let operational_area = row.get(21).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let other_cause = row.get(22).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let other_type = row.get(23).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let state = row.get(24).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let time_of_incident_observed = row.get(25).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let time_of_incident_occurred = row.get(26).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let type_of_operation = row.get(27).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let coordinate_format = row.get(28).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let created_at = row.get(29).and_then(|c| c.get_string()).unwrap_or("").to_string();
            let updated_at = row.get(30).and_then(|c| c.get_string()).unwrap_or("").to_string();

            // Print the extracted data
            println!("Row {}: {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {} {}",
                row_index + 1, id, spill_serial_no, facility_or_equipment, location, incident_date, coordinates, latitude, longitude, spill_status, description_of_spill_causes, estimated_qty_spilled_or_leaked, severity, extent_of_pollution, precaution_measures, client_id, opl_oml_no, date_of_incident_observed, date_of_incident_occurred, incident_cause, incident_type, nearest_town, operational_area, other_cause, other_type, state, time_of_incident_observed, time_of_incident_occurred, type_of_operation, coordinate_format, created_at, updated_at);

            // Write the row to the new workbook
            let data_row = row_index as u32 + 1; // +1 to account for header row
            worksheet.write_string(data_row, 0, &id)?;
            worksheet.write_string(data_row, 1, &spill_serial_no)?;
            worksheet.write_string(data_row, 2, &facility_or_equipment)?;
            worksheet.write_string(data_row, 3, &location)?;
            worksheet.write_string(data_row, 4, &incident_date)?;
            worksheet.write_string(data_row, 5, &coordinates)?;
            worksheet.write_number(data_row, 6, latitude)?;
            worksheet.write_number(data_row, 7, longitude)?;
            worksheet.write_string(data_row, 8, &spill_status)?;
            worksheet.write_string(data_row, 9, &description_of_spill_causes)?;
            worksheet.write_number(data_row, 10, estimated_qty_spilled_or_leaked)?;
            worksheet.write_string(data_row, 11, &severity)?;
            worksheet.write_string(data_row, 12, &extent_of_pollution)?;
            worksheet.write_string(data_row, 13, &precaution_measures)?;
            worksheet.write_string(data_row, 14, &client_id)?;
            worksheet.write_number(data_row, 15, opl_oml_no as f64)?;
            worksheet.write_string(data_row, 16, &date_of_incident_observed)?;
            worksheet.write_string(data_row, 17, &date_of_incident_occurred)?;
            worksheet.write_string(data_row, 18, &incident_cause)?;
            worksheet.write_string(data_row, 19, &incident_type)?;
            worksheet.write_string(data_row, 20, &nearest_town)?;
            worksheet.write_string(data_row, 21, &operational_area)?;
            worksheet.write_string(data_row, 22, &other_cause)?;
            worksheet.write_string(data_row, 23, &other_type)?;
            worksheet.write_string(data_row, 24, &state)?;
            worksheet.write_string(data_row, 25, &time_of_incident_observed)?;
            worksheet.write_string(data_row, 26, &time_of_incident_occurred)?;
            worksheet.write_string(data_row, 27, &type_of_operation)?;
            worksheet.write_string(data_row, 28, &coordinate_format)?;
            worksheet.write_string(data_row, 29, &created_at)?;
            worksheet.write_string(data_row, 30, &updated_at)?;
        }

        // Save the new workbook with an absolute path to ensure we know where it's saved
        let save_path = "c:\\Users\\bunak\\Development\\Rust\\xcel_operations\\loaded_register.xlsx";
        println!("Saving file to: {}", save_path);
        new_workbook.save(save_path)?;
        println!("File save attempt completed");
    } else {
        println!("Error: Worksheet 'Sheet1' not found. Check sheet names.");
        return Err("Worksheet not found".into());
    }

    Ok(())
}
