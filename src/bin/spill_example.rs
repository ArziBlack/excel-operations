use std::error::Error;
use chrono::{Local, NaiveDate, NaiveTime};
use xcel_operations::spill_register::{SpillIncident, add_spill_incident};

fn main() -> Result<(), Box<dyn Error>> {
    // Create a new spill incident
    let incident = SpillIncident {
        reference_number: "SPL-2025-001".to_string(),
        epage_code: Some("EP-25-001".to_string()),
        facility_equipment: "Storage Tank 3".to_string(),
        facility_category: "Storage".to_string(),
        location: "Building 3, Room 201".to_string(),
        closeby_facility: Some("Pump Station 2".to_string()),
        area: "North Zone".to_string(),
        lga: "Industrial Area".to_string(),
        state: "Lagos".to_string(),
        latitude: Some(6.5244),
        longitude: Some(3.3792),
        incident_date: Local::now().date_naive(),
        incident_time: Some(NaiveTime::from_hms_opt(10, 30, 0).unwrap()),
        observed_date: Some(Local::now().date_naive()),
        observed_time: Some(NaiveTime::from_hms_opt(10, 45, 0).unwrap()),
        stopped_date: Some(Local::now().date_naive()),
        contained_date: Some(Local::now().date_naive()),
        contained_time: Some(NaiveTime::from_hms_opt(11, 15, 0).unwrap()),
        oml_opl: Some("OML-123".to_string()),
        spill_cause: "Equipment Failure".to_string(),
        spill_category: "Minor".to_string(),
        cause_description: Some("Valve malfunction caused hydraulic oil leak".to_string()),
        responsibility: "Internal".to_string(),
        quantity_spilled_bbls: 2.5,
        quantity_recovered_bbls: Some(2.0),
        tiered_level: Some("Tier 1".to_string()),
        cleanup_required: Some(true),
        cleanup_method: Some("Absorbent Materials".to_string()),
        cleanup_contractor: Some("Internal Team".to_string()),
        cleanup_completion_status: Some(100.0),
        cleanup_cost_ngn: Some(50000.0),
        remediation_required: Some(false),
        remediation_method: None,
        communities_impacted: None,
        remarks: Some("Spill contained and cleaned up within 30 minutes.".to_string()),
    };

    // Add the incident to the spill register
    add_spill_incident(incident)?;
    
    println!("Spill incident recorded successfully!");

    // Example of adding another incident
    let another_incident = SpillIncident {
        reference_number: "SPL-2025-002".to_string(),
        epage_code: Some("EP-25-002".to_string()),
        facility_equipment: "Pipeline Section B".to_string(),
        facility_category: "Pipeline".to_string(),
        location: "Workshop Area".to_string(),
        closeby_facility: Some("Compressor Station 1".to_string()),
        area: "South Zone".to_string(),
        lga: "Industrial Area".to_string(),
        state: "Lagos".to_string(),
        latitude: Some(6.5100),
        longitude: Some(3.3900),
        incident_date: Local::now().date_naive(),
        incident_time: Some(NaiveTime::from_hms_opt(14, 15, 0).unwrap()),
        observed_date: Some(Local::now().date_naive()),
        observed_time: Some(NaiveTime::from_hms_opt(14, 20, 0).unwrap()),
        stopped_date: Some(Local::now().date_naive()),
        contained_date: Some(Local::now().date_naive()),
        contained_time: Some(NaiveTime::from_hms_opt(15, 0, 0).unwrap()),
        oml_opl: Some("OML-123".to_string()),
        spill_cause: "Corrosion".to_string(),
        spill_category: "Minor".to_string(),
        cause_description: Some("Small leak from corroded pipe joint".to_string()),
        responsibility: "Internal".to_string(),
        quantity_spilled_bbls: 5.0,
        quantity_recovered_bbls: Some(4.0),
        tiered_level: Some("Tier 1".to_string()),
        cleanup_required: Some(true),
        cleanup_method: Some("Mechanical Recovery".to_string()),
        cleanup_contractor: Some("EcoClean Ltd".to_string()),
        cleanup_completion_status: Some(80.0),
        cleanup_cost_ngn: Some(120000.0),
        remediation_required: Some(true),
        remediation_method: Some("Bioremediation".to_string()),
        communities_impacted: None,
        remarks: Some("Cleanup in progress. Area cordoned off.".to_string()),
    };

    // Add the second incident to the spill register
    add_spill_incident(another_incident)?;
    
    println!("Second spill incident recorded successfully!");

    Ok(())
}
