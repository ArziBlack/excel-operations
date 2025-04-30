use std::error::Error;
use chrono::Local;
use xcel_operations::spill_register::{SpillIncident, add_spill_incident};

fn main() -> Result<(), Box<dyn Error>> {
    // Create a new spill incident
    let incident = SpillIncident {
        date_time: Local::now(),
        location: "Building 3, Room 201".to_string(),
        material: "Hydraulic Oil".to_string(),
        quantity_liters: 2.5,
        reported_by: "John Smith".to_string(),
        cleanup_status: "Completed".to_string(),
        notes: Some("Spill contained and cleaned up within 30 minutes.".to_string()),
    };

    // Add the incident to the spill register
    add_spill_incident(incident)?;
    
    println!("Spill incident recorded successfully!");

    // Example of adding another incident
    let another_incident = SpillIncident {
        date_time: Local::now(),
        location: "Workshop Area".to_string(),
        material: "Coolant".to_string(),
        quantity_liters: 5.0,
        reported_by: "Jane Doe".to_string(),
        cleanup_status: "In Progress".to_string(),
        notes: None,
    };

    // Add the second incident to the spill register
    add_spill_incident(another_incident)?;
    
    println!("Second spill incident recorded successfully!");

    Ok(())
}
