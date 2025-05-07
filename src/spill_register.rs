use std::path::Path;
use std::error::Error;
use chrono::{Local, NaiveDate, NaiveTime};
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Format, Color, FormatBorder};

/// Represents a spill incident to be recorded in the register
/// Fields match the structure of the spill_register.xlsx file
pub struct SpillIncident {
    pub reference_number: String,
    pub epage_code: Option<String>,
    pub facility_equipment: String,
    pub facility_category: String,
    pub location: String,
    pub closeby_facility: Option<String>,
    pub area: String,
    pub lga: String,
    pub state: String,
    pub latitude: Option<f64>,
    pub longitude: Option<f64>,
    pub incident_date: NaiveDate,
    pub incident_time: Option<NaiveTime>,
    pub observed_date: Option<NaiveDate>,
    pub observed_time: Option<NaiveTime>,
    pub stopped_date: Option<NaiveDate>,
    pub contained_date: Option<NaiveDate>,
    pub contained_time: Option<NaiveTime>,
    pub oml_opl: Option<String>,
    pub spill_cause: String,
    pub spill_category: String,
    pub cause_description: Option<String>,
    pub responsibility: String,
    pub quantity_spilled_bbls: f64,
    pub quantity_recovered_bbls: Option<f64>,
    pub tiered_level: Option<String>,
    pub cleanup_required: Option<bool>,
    pub cleanup_method: Option<String>,
    pub cleanup_contractor: Option<String>,
    pub cleanup_completion_status: Option<f64>,  // Percentage
    pub cleanup_cost_ngn: Option<f64>,
    pub remediation_required: Option<bool>,
    pub remediation_method: Option<String>,
    pub communities_impacted: Option<String>,
    pub remarks: Option<String>,
}

impl Default for SpillIncident {
    fn default() -> Self {
        Self {
            reference_number: String::new(),
            epage_code: None,
            facility_equipment: String::new(),
            facility_category: String::new(),
            location: String::new(),
            closeby_facility: None,
            area: String::new(),
            lga: String::new(),
            state: String::new(),
            latitude: None,
            longitude: None,
            incident_date: Local::now().date_naive(),
            incident_time: None,
            observed_date: None,
            observed_time: None,
            stopped_date: None,
            contained_date: None,
            contained_time: None,
            oml_opl: None,
            spill_cause: String::new(),
            spill_category: String::new(),
            cause_description: None,
            responsibility: String::new(),
            quantity_spilled_bbls: 0.0,
            quantity_recovered_bbls: None,
            tiered_level: None,
            cleanup_required: None,
            cleanup_method: None,
            cleanup_contractor: None,
            cleanup_completion_status: None,
            cleanup_cost_ngn: None,
            remediation_required: None,
            remediation_method: None,
            communities_impacted: None,
            remarks: None,
        }
    }
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
    let incident_date_str = incident.incident_date.format("%Y-%m-%d").to_string();
    let incident_time_str = incident.incident_time.map_or(String::new(), |t| t.format("%H:%M:%S").to_string());
    
    // Get a mutable reference to the first worksheet
    // We need to use this approach since rust_xlsxwriter's API requires a mutable worksheet
    let worksheet = &mut workbook.worksheets_mut()[0];
    
    // Write the data to the next empty row
    worksheet.write_string(next_row, 0, &incident.reference_number)?;
    
    if let Some(epage) = &incident.epage_code {
        worksheet.write_string(next_row, 1, epage)?;
    }
    
    worksheet.write_string(next_row, 2, &incident.facility_equipment)?;
    worksheet.write_string(next_row, 3, &incident.facility_category)?;
    worksheet.write_string(next_row, 4, &incident.location)?;
    
    if let Some(closeby) = &incident.closeby_facility {
        worksheet.write_string(next_row, 5, closeby)?;
    }
    
    worksheet.write_string(next_row, 6, &incident.area)?;
    worksheet.write_string(next_row, 7, &incident.lga)?;
    worksheet.write_string(next_row, 8, &incident.state)?;
    
    if let Some(lat) = incident.latitude {
        worksheet.write_number(next_row, 9, lat)?;
    }
    
    if let Some(long) = incident.longitude {
        worksheet.write_number(next_row, 10, long)?;
    }
    
    worksheet.write_string(next_row, 11, &incident_date_str)?;
    
    if let Some(observed_date) = incident.observed_date {
        worksheet.write_string(next_row, 12, &observed_date.format("%Y-%m-%d").to_string())?;
    }
    
    if !incident_time_str.is_empty() {
        worksheet.write_string(next_row, 13, &incident_time_str)?;
    }
    
    if let Some(observed_time) = incident.observed_time {
        worksheet.write_string(next_row, 14, &observed_time.format("%H:%M:%S").to_string())?;
    }
    
    if let Some(stopped_date) = incident.stopped_date {
        worksheet.write_string(next_row, 15, &stopped_date.format("%Y-%m-%d").to_string())?;
    }
    
    if let Some(contained_date) = incident.contained_date {
        worksheet.write_string(next_row, 16, &contained_date.format("%Y-%m-%d").to_string())?;
    }
    
    if let Some(contained_time) = incident.contained_time {
        worksheet.write_string(next_row, 17, &contained_time.format("%H:%M:%S").to_string())?;
    }
    
    if let Some(oml_opl) = &incident.oml_opl {
        worksheet.write_string(next_row, 18, oml_opl)?;
    }
    
    worksheet.write_string(next_row, 19, &incident.spill_cause)?;
    worksheet.write_string(next_row, 20, &incident.spill_category)?;
    
    if let Some(desc) = &incident.cause_description {
        worksheet.write_string(next_row, 21, desc)?;
    }
    
    worksheet.write_string(next_row, 22, &incident.responsibility)?;
    worksheet.write_number(next_row, 23, incident.quantity_spilled_bbls)?;
    
    if let Some(recovered) = incident.quantity_recovered_bbls {
        worksheet.write_number(next_row, 24, recovered)?;
    }
    
    if let Some(tier) = &incident.tiered_level {
        worksheet.write_string(next_row, 25, tier)?;
    }
    
    // Skip to cleanup required column (46)
    if let Some(cleanup_required) = incident.cleanup_required {
        worksheet.write_string(next_row, 45, if cleanup_required { "Yes" } else { "No" })?;
    }
    
    if let Some(method) = &incident.cleanup_method {
        worksheet.write_string(next_row, 49, method)?;
    }
    
    if let Some(contractor) = &incident.cleanup_contractor {
        worksheet.write_string(next_row, 51, contractor)?;
    }
    
    if let Some(completion) = incident.cleanup_completion_status {
        worksheet.write_number(next_row, 62, completion)?;
    }
    
    if let Some(cost) = incident.cleanup_cost_ngn {
        worksheet.write_number(next_row, 56, cost)?;
    }
    
    // Skip to remediation required column (69)
    if let Some(remediation_required) = incident.remediation_required {
        worksheet.write_string(next_row, 68, if remediation_required { "Yes" } else { "No" })?;
    }
    
    if let Some(method) = &incident.remediation_method {
        worksheet.write_string(next_row, 69, method)?;
    }
    
    if let Some(communities) = &incident.communities_impacted {
        worksheet.write_string(next_row, 86, communities)?;
    }
    
    if let Some(remarks) = &incident.remarks {
        // Add remarks to a general remarks column if available
        // This might need adjustment based on the actual Excel structure
        worksheet.write_string(next_row, 42, remarks)?;
    }
    
    // Save the workbook
    workbook.save("spill_register.xlsx")?;
    println!("Spill incident added successfully at row {}", next_row + 1);
    
    Ok(())
}

/// Finds the next empty row in the spill register
fn find_next_empty_row(file_path: &Path) -> Result<u32, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;
    
    if let Some(sheet_name) = workbook.sheet_names().get(0).cloned() {
        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            // Get the number of rows in the sheet (add 1 for the next empty row)
            let row_count = range.height() as u32;
            
            // Return the next row index (0-based)
            Ok(row_count)
        } else {
            // If sheet doesn't exist or is empty, start at row 1 (after header)
            Ok(1)
        }
    } else {
        // If no sheets exist, start at row 1 (after header)
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
    if let Some(sheet_name) = original_workbook.sheet_names().get(0).cloned() {
        if let Ok(range) = original_workbook.worksheet_range(&sheet_name) {
            let worksheet = new_workbook.add_worksheet().set_name(&sheet_name)?;
            
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
            for i in 0..95 {
                worksheet.set_column_width(i, 15.0)?;
            }
            
            Ok(new_workbook)
        } else {
            // If the sheet doesn't exist, create a new one
            create_spill_register()
        }
    } else {
        // If no sheets exist, create a new one
        create_spill_register()
    }
}

/// Creates a new spill register with headers
fn create_spill_register() -> Result<Workbook, Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet().set_name("Sheet1")?;
    
    // Create header format
    let header_format = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0x4F81BD))
        .set_font_color(Color::RGB(0xFFFFFF))
        .set_border(FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center);
    
    // Write title
    worksheet.write_string(0, 0, "Incident/Spill Records")?;
    
    // Write headers
    let headers = [
        "Incident/Oil Spill Ref. No.  ", "e-PAGE code", "Facility/ Equipment  ", "  Facility Category  ",
        "Location", "Closeby Facility", "Area  ", "LGA     ", "State      ", "   Lat (N)   ", "   Long (E)   ",
        "Date of Incident/Spill", "Date of Incident/Spill Observed ", "Time of Incident/Spill",
        "Time of Incident/Spill Observed", "Date spill was stopped", "Date spill was contained",
        "Time spill was contained", " OML/OPL ", "Incident/Spill cause       ", "Incident/ Spill Category    ",
        "Description of Incident/Spill causes", "Incident/Spill responsibility", "Est. Qty. Spilled (bbls)",
        "Est. Qty. Recovered(bbls)", "Tiered Level", "Date of survey", "Survey Map Ref. n.",
        "Survey Cost (NGN) x 1000", "Survey Cost (US$ EQ) x 1000", "NNPC Budget Code", "Date of Mobilisation",
        "Date of Repair ", "Contract Ref. No. used for the repair", "Contractor name", "Repair Cost (NGN) x1000",
        "Repair Cost (US$) x 1000", "Repair Cost (US$ EQV.) x 1000", "Days shutin",
        "Estimated Production loss (bbls)", "Loss due to Oil Spill (NGN)", "Loss due to Oil Spill (US$ EQV.)",
        "Remarks costs involved in the repairs", "COMPLETION STATUS % ", " JIV date ", "Is Cleanup Required?",
        "Date of Form A", "Date of Form B", "Date of Form C NUPRC", "Method of Clean Up", "In house / Contractor",
        "Contractor Name", "Contract Ref. No.", "Date of mobilization ", "Date of Completion ",
        "Estimated duration (days)", "Clean Up Cost (NGN) x 1000", "Clean Up Cost (US$ EQV.) x 1000",
        "Loss due to Oil Spill (NGN) x 1000", "Loss due to Oil Spill (US$ EQV.) x 1000",
        "Down time man hours lost (NGN) x 1000", "Down time man hours lost (US$ EQV.) x 1000",
        "Clean Up Completion Status %", "Any Follow up study?", "Any Compensation required?",
        "Any Rehabilitaion Plan required?", "Date of final sampling once clean up completed", "PCI date  ",
        "Is Remediation required? ", "Remediation method", "Contractor name     ", "Contractor ref. No.",
        "Date of mobilization", "Estimated date of completion ", "Down time man hours lost (NGN) x 1000",
        "Down time man hours lost (US$ EQV.) x 1000", "Effective Date of Remediation Completion",
        "Remediation Cost (NGN) x 1000", "Remediation Cost (US$ EQV.) x 1000",
        "Date of final sampling once remediation completed ", "Date of Close Out Request ",
        "Date of NOSDRA Certificate ", "Extent of contaminated area (Ha)", "Cost of Compensation (NGN) x 1000",
        "Cost of Compensation (US$ EQV.) x 1000", "Relief", "Communities impacted", " A ", " B ", " CD ",
        "Date of FORM C NOSDRA (repair/clean up)", "Date of FORM C NOSDRA (remediation)", " CN ", " CO "
    ];
    
    for (idx, header) in headers.iter().enumerate() {
        worksheet.write_with_format(1, idx as u16, *header, &header_format)?;
    }
    
    // Set column widths for better readability
    for i in 0..headers.len() {
        worksheet.set_column_width(i as u16, 15.0)?;
    }
    
    // Save the workbook
    workbook.save("spill_register.xlsx")?;
    println!("Created new spill register file: spill_register.xlsx");
    
    Ok(workbook)
}
