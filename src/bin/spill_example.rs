use std::error::Error;
use chrono::Utc;
use xcel_operations::spill_register::{SpillIncident, add_spill_incident};

fn main() -> Result<(), Box<dyn Error>> {
    // Create a new spill incident
    let incident = SpillIncident {
        id: "SPL-2025-001".to_string(),
        spill_serial_no: "SPL-2025-001".to_string(),
        facility_or_equipment: "Storage Tank 3".to_string(),
        location: "Building 3, Room 201".to_string(),
        incident_date: Utc::now(),
        coordinates: Some("6.5244, 3.3792".to_string()),
        latitude: Some(6.5244),
        longitude: Some(3.3792),
        spill_status: "ONGOING".to_string(),
        description_of_spill_causes: "Valve malfunction caused hydraulic oil leak".to_string(),
        estimated_qty_spilled_or_leaked: 2.5,
        severity: "LOW".to_string(),
        extent_of_pollution: "Contained within facility".to_string(),
        precaution_measures: "Area cordoned off, cleanup team deployed".to_string(),
        client_id: "CLIENT123".to_string(),
        opl_oml_no: Some(123),
        date_of_incident_or_spill_observed: Some(Utc::now()),
        date_of_incident_or_spill_occurred: Some(Utc::now()),
        incident_cause_of_spill_or_leakage: Some("EQUIPMENT_FAILURE".to_string()),
        incident_type_of_spill_category: Some("PRODUCT".to_string()),
        nearest_town: Some("Lagos".to_string()),
        operational_area: Some("LAND".to_string()),
        other_cause_of_spill_or_leak: None,
        other_type_of_spill: None,
        state: Some("Lagos".to_string()),
        time_of_incident_or_spill_observed: Some("10:45".to_string()),
        time_of_incident_or_spill_occurred: Some("10:30".to_string()),
        type_of_operation_at_spill_site: Some("Storage".to_string()),
        coordinate_format: Some("DD".to_string()),
        created_at: Utc::now(),
        updated_at: Utc::now(),
    };

    // Add the incident to the spill register
    add_spill_incident(incident)?;
    
    println!("Spill incident recorded successfully!");

    // Example of adding another incident
    let another_incident = SpillIncident {
        id: "SPL-2025-002".to_string(),
        spill_serial_no: "SPL-2025-002".to_string(),
        facility_or_equipment: "Pipeline Section B".to_string(),
        location: "Workshop Area".to_string(),
        incident_date: Utc::now(),
        coordinates: Some("6.5100, 3.3900".to_string()),
        latitude: Some(6.5100),
        longitude: Some(3.3900),
        spill_status: "ONGOING".to_string(),
        description_of_spill_causes: "Small leak from corroded pipe joint".to_string(),
        estimated_qty_spilled_or_leaked: 5.0,
        severity: "LOW".to_string(),
        extent_of_pollution: "Contamination spread to nearby area".to_string(),
        precaution_measures: "Area cordoned off, mechanical recovery in progress".to_string(),
        client_id: "CLIENT123".to_string(),
        opl_oml_no: Some(123),
        date_of_incident_or_spill_observed: Some(Utc::now()),
        date_of_incident_or_spill_occurred: Some(Utc::now()),
        incident_cause_of_spill_or_leakage: Some("CORROSION".to_string()),
        incident_type_of_spill_category: Some("PRODUCT".to_string()),
        nearest_town: Some("Lagos".to_string()),
        operational_area: Some("LAND".to_string()),
        other_cause_of_spill_or_leak: None,
        other_type_of_spill: None,
        state: Some("Lagos".to_string()),
        time_of_incident_or_spill_observed: Some("14:20".to_string()),
        time_of_incident_or_spill_occurred: Some("14:15".to_string()),
        type_of_operation_at_spill_site: Some("Pipeline".to_string()),
        coordinate_format: Some("DD".to_string()),
        created_at: Utc::now(),
        updated_at: Utc::now(),
    };

    // Add the second incident to the spill register
    add_spill_incident(another_incident)?;
    
    println!("Second spill incident recorded successfully!");

    Ok(())
}
