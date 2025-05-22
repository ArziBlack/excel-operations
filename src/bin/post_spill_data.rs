use calamine::{open_workbook, Reader, Xlsx};
use reqwest::Client;
use serde_json::json;
use std::error::Error;
use tokio;

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    // Open the Excel file
    let path = "c:\\Users\\bunak\\Development\\Rust\\xcel_operations\\Oil_Spill_2013.xlsx";
    println!("Opening file: {}", path);
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    println!("Workbook opened successfully");

    // Print all sheet names for debugging
    let sheet_names = workbook.sheet_names();
    println!("Available sheets: {:?}", sheet_names);

    // Select a worksheet by name
    if let Ok(range) = workbook.worksheet_range("Oil_Spill_2013") {
        println!("Worksheet selected with {} rows", range.rows().count());

        // Read headers from the first row
        let dynamic_headers: Vec<String> = range.rows().next().unwrap_or(&[]).iter().map(|cell| cell.to_string()).collect();
        println!("Headers read: {:?}", dynamic_headers);

        // Define expected field names from OilSpillIncident struct
        let expected_fields: Vec<&'static str> = vec![
            "incident_oil_spill_ref_no", "e_page_code", "facility_equipment", "facility_category", 
            "location", "closeby_facility", "area", "lga", "state", "coordinates_format", 
            "latitude", "longitude", "date_of_incident_spill", "date_of_incident_spill_observed", 
            "time_of_incident_spill", "time_of_incident_spill_observed", "date_spill_was_stopped", 
            "date_spill_was_contained", "time_spill_was_contained", "oml_opl", 
            "incident_cause_of_spill_or_leakage", "incident_type_spill_category", 
            "description_of_spill_causes", "incident_spill_responsibility", "est_qty_spilled", 
            "est_qty_recovered", "tiered_level", "clean_up_status", "clean_up_completion_date", 
            "clean_up_completion_time", "clean_up_method", "post_clean_up_status", 
            "evidence_of_clean_up", "evidence_of_clean_up_date", "evidence_of_clean_up_time", 
            "post_clean_up_observation", "post_incident_impact", "post_incident_impact_description", 
            "incident_investigation", "incident_investigation_date", "incident_investigation_time", 
            "incident_investigation_team", "incident_investigation_team_leader", 
            "incident_investigation_team_leader_phone", "incident_investigation_team_leader_email", 
            "incident_investigation_team_members", "client_id", "client", "client_phone", 
            "client_email", "client_address", "client_representative", 
            "client_representative_phone", "client_representative_email", "client_representative_role", 
            "client_representative_office_address", "contractor", "contractor_phone", 
            "contractor_email", "contractor_address", "contractor_representative", 
            "contractor_representative_phone", "contractor_representative_email", 
            "contractor_representative_role", "contractor_representative_office_address", 
            "financial_cost", "financial_currency", "financial_description", 
            "financial_payment_status", "financial_payment_date", "financial_payment_time", 
            "financial_payment_receipt", "financial_payment_receipt_date", 
            "financial_payment_receipt_time"
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
                    (header_lower.contains("ref. no.") && f_lower == "incident_oil_spill_ref_no") ||
                    (header_lower.contains("spill ref no") && f_lower == "incident_oil_spill_ref_no") ||
                    (header_lower.contains("lat") && f_lower == "latitude") ||
                    (header_lower.contains("long") && f_lower == "longitude") ||
                    (header_lower.contains("cause") && f_lower == "incident_cause_of_spill_or_leakage") ||
                    (header_lower.contains("location") && f_lower == "location") ||
                    (header_lower.contains("state") && f_lower == "state") ||
                    (header_lower.contains("facility") && f_lower == "facility_equipment") ||
                    (header_lower.contains("category") && f_lower == "facility_category") ||
                    (header_lower.contains("closeby") && f_lower == "closeby_facility") ||
                    (header_lower.contains("area") && f_lower == "area") ||
                    (header_lower.contains("lga") && f_lower == "lga") ||
                    (header_lower.contains("coordinate") && f_lower == "coordinates_format") ||
                    (header_lower.contains("incident date") && f_lower == "date_of_incident_spill") ||
                    (header_lower.contains("observed date") && f_lower == "date_of_incident_spill_observed") ||
                    (header_lower.contains("incident time") && f_lower == "time_of_incident_spill") ||
                    (header_lower.contains("observed time") && f_lower == "time_of_incident_spill_observed") ||
                    (header_lower.contains("stopped date") && f_lower == "date_spill_was_stopped") ||
                    (header_lower.contains("date stopped") && f_lower == "date_spill_was_stopped") ||
                    (header_lower.contains("contained date") && f_lower == "date_spill_was_contained") ||
                    (header_lower.contains("date contained") && f_lower == "date_spill_was_contained") ||
                    (header_lower.contains("contained time") && f_lower == "time_spill_was_contained") ||
                    (header_lower.contains("time contained") && f_lower == "time_spill_was_contained") ||
                    (header_lower.contains("oml/opl") && f_lower == "oml_opl") ||
                    (header_lower.contains("spill type") && f_lower == "incident_type_spill_category") ||
                    (header_lower.contains("spill description") && f_lower == "description_of_spill_causes") ||
                    (header_lower.contains("responsibility") && f_lower == "incident_spill_responsibility") ||
                    (header_lower.contains("spilled qty") && f_lower == "est_qty_spilled") ||
                    (header_lower.contains("recovered qty") && f_lower == "est_qty_recovered") ||
                    (header_lower.contains("tier") && f_lower == "tiered_level") ||
                    (header_lower.contains("clean up") && f_lower == "clean_up_status") ||
                    (header_lower.contains("completion date") && f_lower == "clean_up_completion_date") ||
                    (header_lower.contains("completion time") && f_lower == "clean_up_completion_time") ||
                    (header_lower.contains("method") && f_lower == "clean_up_method") ||
                    (header_lower.contains("post clean") && f_lower == "post_clean_up_status") ||
                    (header_lower.contains("evidence") && f_lower == "evidence_of_clean_up") ||
                    (header_lower.contains("evidence date") && f_lower == "evidence_of_clean_up_date") ||
                    (header_lower.contains("evidence time") && f_lower == "evidence_of_clean_up_time") ||
                    (header_lower.contains("observation") && f_lower == "post_clean_up_observation") ||
                    (header_lower.contains("impact") && f_lower == "post_incident_impact") ||
                    (header_lower.contains("impact description") && f_lower == "post_incident_impact_description") ||
                    (header_lower.contains("investigation") && f_lower == "incident_investigation") ||
                    (header_lower.contains("investigation date") && f_lower == "incident_investigation_date") ||
                    (header_lower.contains("investigation time") && f_lower == "incident_investigation_time") ||
                    (header_lower.contains("team") && f_lower == "incident_investigation_team") ||
                    (header_lower.contains("team leader") && f_lower == "incident_investigation_team_leader") ||
                    (header_lower.contains("leader phone") && f_lower == "incident_investigation_team_leader_phone") ||
                    (header_lower.contains("leader email") && f_lower == "incident_investigation_team_leader_email") ||
                    (header_lower.contains("team members") && f_lower == "incident_investigation_team_members") ||
                    (header_lower.contains("client id") && f_lower == "client_id") ||
                    (header_lower.contains("client") && f_lower == "client") ||
                    (header_lower.contains("client phone") && f_lower == "client_phone") ||
                    (header_lower.contains("client email") && f_lower == "client_email") ||
                    (header_lower.contains("client address") && f_lower == "client_address") ||
                    (header_lower.contains("representative") && f_lower == "client_representative") ||
                    (header_lower.contains("representative phone") && f_lower == "client_representative_phone") ||
                    (header_lower.contains("representative email") && f_lower == "client_representative_email") ||
                    (header_lower.contains("representative role") && f_lower == "client_representative_role") ||
                    (header_lower.contains("office address") && f_lower == "client_representative_office_address") ||
                    (header_lower.contains("contractor") && f_lower == "contractor") ||
                    (header_lower.contains("contractor phone") && f_lower == "contractor_phone") ||
                    (header_lower.contains("contractor email") && f_lower == "contractor_email") ||
                    (header_lower.contains("contractor address") && f_lower == "contractor_address") ||
                    (header_lower.contains("contractor rep") && f_lower == "contractor_representative") ||
                    (header_lower.contains("contractor rep phone") && f_lower == "contractor_representative_phone") ||
                    (header_lower.contains("contractor rep email") && f_lower == "contractor_representative_email") ||
                    (header_lower.contains("contractor rep role") && f_lower == "contractor_representative_role") ||
                    (header_lower.contains("contractor office") && f_lower == "contractor_representative_office_address") ||
                    (header_lower.contains("cost") && f_lower == "financial_cost") ||
                    (header_lower.contains("currency") && f_lower == "financial_currency") ||
                    (header_lower.contains("financial description") && f_lower == "financial_description") ||
                    (header_lower.contains("payment status") && f_lower == "financial_payment_status") ||
                    (header_lower.contains("payment date") && f_lower == "financial_payment_date") ||
                    (header_lower.contains("payment time") && f_lower == "financial_payment_time") ||
                    (header_lower.contains("payment receipt") && f_lower == "financial_payment_receipt") ||
                    (header_lower.contains("receipt date") && f_lower == "financial_payment_receipt_date") ||
                    (header_lower.contains("receipt time") && f_lower == "financial_payment_receipt_time") ||
                    header_lower.contains(&f_lower)
                })
            })
            .map(|(i, header)| (i, header.clone()))
            .collect();
        println!("Matching headers and indices: {:?}", matching_indices);

        // Initialize HTTP client
        let client = Client::new();
        let endpoint = "https://f3d-server.onrender.com/api/v1/spills";

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

            // Create JSON object for the row
            let mut data_obj = json!({});
            for (i, data) in row_data.iter().enumerate() {
                let field_name = &matching_indices[i].1;
                data_obj[field_name] = json!(data);
            }

            // Send data to endpoint
            match client.post(endpoint)
                .json(&data_obj)
                .send()
                .await {
                Ok(response) => println!("Row {} sent successfully: {:?}", row_index, response.status()),
                Err(e) => println!("Error sending row {}: {}", row_index, e),
            }
        }
    } else {
        println!("Error: Worksheet 'Oil_Spill_2013' not found. Check sheet names.");
        return Err("Worksheet not found".into());
    }

    Ok(())
}
