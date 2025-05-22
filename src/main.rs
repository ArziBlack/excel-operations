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

    // Load and process the spill data dynamically
    println!("\nExample 4: Loading and processing spill data dynamically");
    load_and_process_spill_data_dynamic()?;

    println!("\nAll examples completed successfully!");
    println!("Files created:");
    println!("- data.xlsx (sample data)");
    println!("- output.xlsx (from write example)");
    println!("- modified_data.xlsx (from modify example)");
    println!("- loaded_register_dynamic.xlsx (from load and process spill data dynamically)");

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

fn load_and_process_spill_data_dynamic() -> Result<(), Box<dyn Error>> {
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

        // Read headers from the first row
        let dynamic_headers: Vec<String> = range.rows().next().unwrap_or(&[]).iter().map(|cell| cell.to_string()).collect();
        println!("Headers read: {:?}", dynamic_headers);

        // Define expected field names from SpillIncident (adjust these based on your struct)
        let expected_fields = vec![
            "ID", "Spill Serial No", "Facility or Equipment", "Location", "Incident Date", 
            "Coordinates", "Latitude", "Longitude", "Spill Status", "Description of Spill Causes", 
            "Estimated Qty Spilled or Leaked", "Severity", "Extent of Pollution", "Precaution Measures", 
            "Client ID", "OPL/OML No", "Date of Incident Observed", "Date of Incident Occurred", 
            "Incident Cause", "Incident Type", "Nearest Town", "Operational Area", "Other Cause", 
            "Other Type", "State", "Time of Incident Observed", "Time of Incident Occurred", 
            "Type of Operation", "Coordinate Format", "Created At", "Updated At"
        ];

        // Find indices of matching headers (case-insensitive)
        let matching_indices: Vec<(usize, String)> = dynamic_headers.iter().enumerate()
            .filter(|(_, header)| expected_fields.iter().any(|f| f.to_lowercase() == header.to_lowercase()))
            .map(|(i, header)| (i, header.clone()))
            .collect();
        println!("Matching headers and indices: {:?}", matching_indices);

        // Write headers for matching fields in new workbook
        for (col_idx, (_, header)) in matching_indices.iter().enumerate() {
            worksheet.write_string(0, col_idx as u16, header)?;
        }

        // Iterate through data rows and extract only matching fields
        for (row_index, row) in range.rows().enumerate().skip(1) { // Skip header row
            println!("Processing row {} with {} cells", row_index, row.len());
            let mut row_data = vec![];
            for &(orig_index, _) in &matching_indices {
                if let Some(cell) = row.get(orig_index) {
                    row_data.push(cell.to_string());
                } else {
                    row_data.push("".to_string());
                }
            }
            println!("Extracted data for row {}: {:?}", row_index, row_data);

            // Write data to new workbook
            for (col_idx, data) in row_data.iter().enumerate() {
                if let Ok(num_val) = data.parse::<f64>() {
                    worksheet.write_number(row_index as u32, col_idx as u16, num_val)?;
                } else {
                    worksheet.write_string(row_index as u32, col_idx as u16, data)?;
                }
            }
        }

        // Save the new workbook with an absolute path to ensure we know where it's saved
        let save_path = "c:\\Users\\bunak\\Development\\Rust\\xcel_operations\\loaded_register_dynamic.xlsx";
        println!("Saving file to: {}", save_path);
        new_workbook.save(save_path)?;
        println!("File save attempt completed");
    } else {
        println!("Error: Worksheet 'Oil_Spill_2013' not found. Check sheet names.");
        return Err("Worksheet not found".into());
    }

    Ok(())
}
