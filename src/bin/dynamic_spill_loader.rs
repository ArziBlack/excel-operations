use std::path::Path;
use std::error::Error;
use std::env;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    // Get command line arguments
    let args: Vec<String> = env::args().collect();
    
    // Default field names to match with struct if no arguments provided
    let expected_fields: Vec<String> = if args.len() > 1 {
        args[1..].to_vec()
    } else {
        vec![
            "Incident/Oil Spill Ref. No.".to_string(),
            "e-PAGE code".to_string(),
            "Facility/ Equipment".to_string(),
            "Facility Category".to_string(),
            "Location".to_string(),
            "Area".to_string(),
            "Coordinates".to_string(),
            "Lat (N)".to_string(),
            "Long (E)".to_string(),
            "Spill Status".to_string(),
            "Description of Spill Causes".to_string(),
            "Estimated Qty Spilled or Leaked".to_string(),
            "LGA".to_string(),
            "Extent of Pollution".to_string(),
            "Precaution Measures".to_string(),
            "State".to_string(),
            "OPL/OML No".to_string(),
            "Date of Incident Observed".to_string(),
            "Date of Incident Occurred".to_string(),
            "Incident Cause".to_string(),
            "Incident Type".to_string(),
            "Nearest Town".to_string(),
            "Operational Area".to_string(),
            "Other Cause".to_string(),
            "Other Type".to_string(),
            "State".to_string(),
            "Time of Incident Observed".to_string(),
            "Time of Incident Occurred".to_string(),
            "Type of Operation".to_string(),
            "Coordinate Format".to_string(),
            "Created At".to_string(),
            "Updated At".to_string()
        ]
    };
    
    println!("Using field names: {:?}", expected_fields);
    load_and_process_spill_data_dynamic(expected_fields)
}

fn load_and_process_spill_data_dynamic(expected_fields: Vec<String>) -> Result<(), Box<dyn Error>> {
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

        // Find indices of matching headers (case-insensitive) with synonym support
        let matching_indices: Vec<(usize, String)> = dynamic_headers.iter().enumerate()
            .filter(|(_, header)| {
                let header_lower = header.trim().to_lowercase();
                expected_fields.iter().any(|f| {
                    let f_lower = f.to_lowercase();
                    // Direct match
                    f_lower == header_lower ||
                    // Synonym or partial match examples (expand this list as needed)
                    (f_lower == "latitude" && header_lower.contains("lat")) ||
                    (f_lower == "longitude" && header_lower.contains("long")) ||
                    (f_lower == "incident_cause_of_spill_or_leakage" && header_lower.contains("cause")) ||
                    (f_lower == "location" && header_lower.contains("location")) ||
                    (f_lower == "state" && header_lower.contains("state"))
                })
            })
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
