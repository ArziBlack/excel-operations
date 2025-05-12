use serde::{Deserialize, Serialize};
use chrono::{DateTime, Utc};
use std::error::Error;
use std::path::Path;
use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Format, Color, FormatBorder};

#[derive(Debug, Serialize, Deserialize)]
pub struct SpillIncident {
    #[serde(rename = "_id")]
    pub id: String,
    pub spill_serial_no: String,
    pub facility_or_equipment: String,
    pub location: String,
    #[serde(rename = "date_of_incident_or_spill")]
    pub incident_date: DateTime<Utc>,
    #[serde(default)]
    pub coordinates: Option<String>,
    #[serde(rename = "lat")]
    pub latitude: Option<f64>,
    #[serde(rename = "long")]
    pub longitude: Option<f64>,
    pub spill_status: String,
    pub description_of_spill_causes: String,
    pub estimated_qty_spilled_or_leaked: f64,
    pub severity: String,
    pub extent_of_pollution: String,
    pub precaution_measures: String,
    pub client_id: String,
    #[serde(rename = "OPL_OML_no_unit_desc")]
    pub opl_oml_no: Option<i32>,
    pub date_of_incident_or_spill_observed: Option<DateTime<Utc>>,
    pub date_of_incident_or_spill_occurred: Option<DateTime<Utc>>,
    pub incident_cause_of_spill_or_leakage: Option<String>,
    pub incident_type_of_spill_category: Option<String>,
    pub nearest_town: Option<String>,
    pub operational_area: Option<String>,
    pub other_cause_of_spill_or_leak: Option<String>,
    pub other_type_of_spill: Option<String>,
    pub state: Option<String>,
    pub time_of_incident_or_spill_observed: Option<String>,
    pub time_of_incident_or_spill_occurred: Option<String>,
    pub type_of_operation_at_spill_site: Option<String>,
    pub coordinate_format: Option<String>,
    #[serde(rename = "createdAt")]
    pub created_at: DateTime<Utc>,
    #[serde(rename = "updatedAt")]
    pub updated_at: DateTime<Utc>,
}

impl Default for SpillIncident {
    fn default() -> Self {
        Self {
            id: String::new(),
            spill_serial_no: String::new(),
            facility_or_equipment: String::new(),
            location: String::new(),
            incident_date: Utc::now(),
            coordinates: None,
            latitude: None,
            longitude: None,
            spill_status: "ONGOING".to_string(),
            description_of_spill_causes: String::new(),
            estimated_qty_spilled_or_leaked: 0.0,
            severity: "LOW".to_string(),
            extent_of_pollution: String::new(),
            precaution_measures: String::new(),
            client_id: String::new(),
            opl_oml_no: None,
            date_of_incident_or_spill_observed: None,
            date_of_incident_or_spill_occurred: None,
            incident_cause_of_spill_or_leakage: None,
            incident_type_of_spill_category: None,
            nearest_town: None,
            operational_area: None,
            other_cause_of_spill_or_leak: None,
            other_type_of_spill: None,
            state: None,
            time_of_incident_or_spill_observed: None,
            time_of_incident_or_spill_occurred: None,
            type_of_operation_at_spill_site: None,
            coordinate_format: None,
            created_at: Utc::now(),
            updated_at: Utc::now(),
        }
    }
}

/// Adds a new spill incident to the spill register
pub fn add_spill_incident(incident: SpillIncident) -> Result<(), Box<dyn Error>> {
    let file_path = std::env::current_dir()?.join("spill_register.xlsx");

     // Check if file exists, if not create it with headers
    if !file_path.exists() {
        create_spill_register()?;
    }
    
    // First read the existing file to find the next empty row
    let next_row = find_next_empty_row(&file_path)?;
    
    // Now open the file for writing
    let mut workbook = open_existing_workbook(&file_path)?;
    
    // Format the date as string (since we can't use DateTime directly)
    let incident_date_str = incident.incident_date.format("%Y-%m-%d").to_string();
    
    // Get a mutable reference to the first worksheet
    let worksheet = &mut workbook.worksheets_mut()[0];
    
    // Write the data to the next empty row
    worksheet.write_string(next_row, 0, &incident.spill_serial_no)?;
    worksheet.write_string(next_row, 2, &incident.facility_or_equipment)?;
    worksheet.write_string(next_row, 4, &incident.location)?;
    
    worksheet.write_string(next_row, 11, &incident_date_str)?;
    
    if let Some(state) = &incident.state {
        worksheet.write_string(next_row, 8, state)?;
    }
    
    if let Some(lat) = incident.latitude {
        worksheet.write_number(next_row, 9, lat)?;
    }
    
    if let Some(long) = incident.longitude {
        worksheet.write_number(next_row, 10, long)?;
    }
    
    if let Some(spill_type) = &incident.incident_type_of_spill_category {
        worksheet.write_string(next_row, 19, spill_type)?;
    }
    
    if let Some(spill_cause) = &incident.incident_cause_of_spill_or_leakage {
        worksheet.write_string(next_row, 20, spill_cause)?;
    }
    
    worksheet.write_number(next_row, 23, incident.estimated_qty_spilled_or_leaked)?;
    
    // For now, we'll use the same value for recovered quantity since it's not available
    worksheet.write_number(next_row, 24, incident.estimated_qty_spilled_or_leaked)?;
    
    worksheet.write_string(next_row, 62, &incident.spill_status)?;
    
    worksheet.write_string(next_row, 42, &incident.description_of_spill_causes)?;
    
    // Save the workbook
    workbook.save(file_path)?;

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
