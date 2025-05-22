use std::path::Path;
use std::error::Error;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    load_and_process_spill_data_dynamic()
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

        // Define expected field names from OilSpillIncident struct
        let expected_fields: Vec<&'static str> = vec![
            "Incident Oil Spill Ref No", "E Page Code", "Facility Equipment", "Facility Category", 
            "Location", "Closeby Facility", "Area", "LGA", "State", "Coordinates Format", 
            "Latitude", "Longitude", "Date of Incident Spill", "Date of Incident Spill Observed", 
            "Time of Incident Spill", "Time of Incident Spill Observed", "Date Spill was Stopped", 
            "Date Spill was Contained", "Time Spill was Contained", "OML OPL", 
            "Incident Cause of Spill or Leakage", "Incident Type Spill Category", 
            "Description of Spill Causes", "Incident Spill Responsibility", "Est Qty Spilled", 
            "Est Qty Recovered", "Tiered Level", "Clean Up Status", "Clean Up Completion Date", 
            "Clean Up Completion Time", "Clean Up Method", "Post Clean Up Status", 
            "Evidence of Clean Up", "Evidence of Clean Up Date", "Evidence of Clean Up Time", 
            "Post Clean Up Observation", "Post Incident Impact", "Post Incident Impact Description", 
            "Incident Investigation", "Incident Investigation Date", "Incident Investigation Time", 
            "Incident Investigation Team", "Incident Investigation Team Leader", 
            "Incident Investigation Team Leader Phone", "Incident Investigation Team Leader Email", 
            "Incident Investigation Team Members", "Client ID", "Client", "Client Phone", 
            "Client Email", "Client Address", "Client Representative", 
            "Client Representative Phone", "Client Representative Email", "Client Representative Role", 
            "Client Representative Office Address", "Contractor", "Contractor Phone", 
            "Contractor Email", "Contractor Address", "Contractor Representative", 
            "Contractor Representative Phone", "Contractor Representative Email", 
            "Contractor Representative Role", "Contractor Representative Office Address", 
            "Financial Cost", "Financial Currency", "Financial Description", 
            "Financial Payment Status", "Financial Payment Date", "Financial Payment Time", 
            "Financial Payment Receipt", "Financial Payment Receipt Date", 
            "Financial Payment Receipt Time"
        ];

        // Find indices of matching headers (case-insensitive) with synonym support
        let matching_indices: Vec<(usize, String)> = dynamic_headers.iter().enumerate()
            .filter(|(_, header)| {
                let header_lower = header.trim().to_lowercase();
                expected_fields.iter().any(|f| {
                    let f_lower = f.to_lowercase();
                    // Direct match
                    f_lower == header_lower ||
                    // Synonym or partial match examples (expand this list as needed)
                    (f_lower == "incident oil spill ref no" && header_lower.contains("ref. no.")) ||
                    (f_lower == "incident oil spill ref no" && header_lower.contains("spill ref no")) ||
                    (f_lower == "latitude" && header_lower.contains("lat")) ||
                    (f_lower == "longitude" && header_lower.contains("long")) ||
                    (f_lower == "incident cause of spill or leakage" && header_lower.contains("cause")) ||
                    (f_lower == "location" && header_lower.contains("location")) ||
                    (f_lower == "state" && header_lower.contains("state")) ||
                    (f_lower == "facility equipment" && header_lower.contains("facility")) ||
                    (f_lower == "facility category" && header_lower.contains("category")) ||
                    (f_lower == "closeby facility" && header_lower.contains("closeby")) ||
                    (f_lower == "area" && header_lower.contains("area")) ||
                    (f_lower == "lga" && header_lower.contains("lga")) ||
                    (f_lower == "coordinates format" && header_lower.contains("coordinate")) ||
                    (f_lower == "date of incident spill" && header_lower.contains("incident date")) ||
                    (f_lower == "date of incident spill observed" && header_lower.contains("observed date")) ||
                    (f_lower == "time of incident spill" && header_lower.contains("incident time")) ||
                    (f_lower == "time of incident spill observed" && header_lower.contains("observed time")) ||
                    (f_lower == "date spill was stopped" && header_lower.contains("stopped date")) ||
                    (f_lower == "date spill was stopped" && header_lower.contains("date stopped")) ||
                    (f_lower == "date spill was contained" && header_lower.contains("contained date")) ||
                    (f_lower == "date spill was contained" && header_lower.contains("date contained")) ||
                    (f_lower == "time spill was contained" && header_lower.contains("contained time")) ||
                    (f_lower == "time spill was contained" && header_lower.contains("time contained")) ||
                    (f_lower == "oml opl" && header_lower.contains("oml/opl")) ||
                    (f_lower == "incident type spill category" && header_lower.contains("spill type")) ||
                    (f_lower == "description of spill causes" && header_lower.contains("spill description")) ||
                    (f_lower == "incident spill responsibility" && header_lower.contains("responsibility")) ||
                    (f_lower == "est qty spilled" && header_lower.contains("spilled qty")) ||
                    (f_lower == "est qty recovered" && header_lower.contains("recovered qty")) ||
                    (f_lower == "tiered level" && header_lower.contains("tier")) ||
                    (f_lower == "clean up status" && header_lower.contains("clean up")) ||
                    (f_lower == "clean up completion date" && header_lower.contains("completion date")) ||
                    (f_lower == "clean up completion time" && header_lower.contains("completion time")) ||
                    (f_lower == "clean up method" && header_lower.contains("method")) ||
                    (f_lower == "post clean up status" && header_lower.contains("post clean")) ||
                    (f_lower == "evidence of clean up" && header_lower.contains("evidence")) ||
                    (f_lower == "evidence of clean up date" && header_lower.contains("evidence date")) ||
                    (f_lower == "evidence of clean up time" && header_lower.contains("evidence time")) ||
                    (f_lower == "post clean up observation" && header_lower.contains("observation")) ||
                    (f_lower == "post incident impact" && header_lower.contains("impact")) ||
                    (f_lower == "post incident impact description" && header_lower.contains("impact description")) ||
                    (f_lower == "incident investigation" && header_lower.contains("investigation")) ||
                    (f_lower == "incident investigation date" && header_lower.contains("investigation date")) ||
                    (f_lower == "incident investigation time" && header_lower.contains("investigation time")) ||
                    (f_lower == "incident investigation team" && header_lower.contains("team")) ||
                    (f_lower == "incident investigation team leader" && header_lower.contains("team leader")) ||
                    (f_lower == "incident investigation team leader phone" && header_lower.contains("leader phone")) ||
                    (f_lower == "incident investigation team leader email" && header_lower.contains("leader email")) ||
                    (f_lower == "incident investigation team members" && header_lower.contains("team members")) ||
                    (f_lower == "client id" && header_lower.contains("client id")) ||
                    (f_lower == "client" && header_lower.contains("client")) ||
                    (f_lower == "client phone" && header_lower.contains("client phone")) ||
                    (f_lower == "client email" && header_lower.contains("client email")) ||
                    (f_lower == "client address" && header_lower.contains("client address")) ||
                    (f_lower == "client representative" && header_lower.contains("representative")) ||
                    (f_lower == "client representative phone" && header_lower.contains("representative phone")) ||
                    (f_lower == "client representative email" && header_lower.contains("representative email")) ||
                    (f_lower == "client representative role" && header_lower.contains("representative role")) ||
                    (f_lower == "client representative office address" && header_lower.contains("office address")) ||
                    (f_lower == "contractor" && header_lower.contains("contractor")) ||
                    (f_lower == "contractor phone" && header_lower.contains("contractor phone")) ||
                    (f_lower == "contractor email" && header_lower.contains("contractor email")) ||
                    (f_lower == "contractor address" && header_lower.contains("contractor address")) ||
                    (f_lower == "contractor representative" && header_lower.contains("contractor rep")) ||
                    (f_lower == "contractor representative phone" && header_lower.contains("contractor rep phone")) ||
                    (f_lower == "contractor representative email" && header_lower.contains("contractor rep email")) ||
                    (f_lower == "contractor representative role" && header_lower.contains("contractor rep role")) ||
                    (f_lower == "contractor representative office address" && header_lower.contains("contractor office")) ||
                    (f_lower == "financial cost" && header_lower.contains("cost")) ||
                    (f_lower == "financial currency" && header_lower.contains("currency")) ||
                    (f_lower == "financial description" && header_lower.contains("financial description")) ||
                    (f_lower == "financial payment status" && header_lower.contains("payment status")) ||
                    (f_lower == "financial payment date" && header_lower.contains("payment date")) ||
                    (f_lower == "financial payment time" && header_lower.contains("payment time")) ||
                    (f_lower == "financial payment receipt" && header_lower.contains("payment receipt")) ||
                    (f_lower == "financial payment receipt date" && header_lower.contains("receipt date")) ||
                    (f_lower == "financial payment receipt time" && header_lower.contains("receipt time"))
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
        let save_path = "c:\\Users\\bunak\\Development\\Rust\\xcel_operations\\loaded_register_dynamic_processed.xlsx";
        println!("Saving file to: {}", save_path);
        new_workbook.save(save_path)?;
        println!("File save attempt completed");
    } else {
        println!("Error: Worksheet 'Oil_Spill_2013' not found. Check sheet names.");
        return Err("Worksheet not found".into());
    }

    Ok(())
}
