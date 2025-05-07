use crate::spill_register::SpillIncident;
use reqwest::Client;
use rust_xlsxwriter::{Format, Workbook, Color, FormatBorder};
use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize)]
pub struct SpillResponse {
    pub status: i32,
    pub success: bool,
    pub message: String,
    pub data: Vec<SpillIncident>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ErrorResponse {
    pub success: bool,
    pub message: String,
    #[serde(default)]
    pub error: Option<serde_json::Value>,  // Optional error data
}

#[derive(Debug)]
pub enum ApiError {
    RequestFailed(ErrorResponse),
    NetworkError(reqwest::Error),
    Other(Box<dyn std::error::Error>),
}

impl std::fmt::Display for ApiError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ApiError::RequestFailed(err) => write!(f, "API request failed: {}", err.message),
            ApiError::NetworkError(err) => write!(f, "Network error: {}", err),
            ApiError::Other(err) => write!(f, "Error: {}", err),
        }
    }
}

impl std::error::Error for ApiError {}

pub async fn fetch_spills(client_id: &str, token: &str) -> Result<SpillResponse, ApiError> {
    let client = Client::new();
    let response = client
        .get("https://f3d-server.onrender.com/api/v1/spills")
        .query(&[("client_id", client_id)])
        .bearer_auth(token)
        .send()
        .await
        .map_err(ApiError::NetworkError)?;

    let status = response.status();
    if status.is_success() {
        // Get the response text first to debug
        let text = response.text().await.map_err(ApiError::NetworkError)?;
        println!("API Response: {}", text);
        
        // Parse the text back to JSON
        let parsed: SpillResponse = serde_json::from_str(&text)
            .map_err(|e| ApiError::Other(Box::new(e)))?;
        Ok(parsed)
    } else {
        let error_response = response
            .json::<ErrorResponse>()
            .await
            .map_err(|e| ApiError::NetworkError(e))?;
        Err(ApiError::RequestFailed(error_response))
    }
}

pub fn save_spills_to_excel(spills: &[SpillIncident], filename: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write headers
    let headers = [
        "ID",
        "Spill Serial No",
        "Facility/Equipment",
        "Location",
        "Incident Date",
        "Coordinates",
        "Latitude",
        "Longitude",
        "Spill Status",
        "Description of Spill Causes",
        "Estimated Qty Spilled/Leaked",
        "Severity",
        "Extent of Pollution",
        "Precaution Measures",
        "Client ID",
        "OPL/OML No",
        "Date Observed",
        "Date Occurred",
        "Incident Cause",
        "Spill Category",
        "Nearest Town",
        "Operational Area",
        "Other Cause",
        "Other Type",
        "State",
        "Time Observed",
        "Time Occurred",
        "Operation Type at Site",
        "Coordinate Format",
        "Created At",
        "Updated At"
    ];

    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::Blue)
        .set_border(FormatBorder::Thin);

    for (col, header) in headers.iter().enumerate() {
        worksheet.write_string_with_format(0, col as u16, *header, &header_format)?;
    }

    // Write data
    for (row, spill) in spills.iter().enumerate() {
        let row = (row + 1) as u32;

        // Write each field
        worksheet.write_string(row, 0, &spill.id)?;
        worksheet.write_string(row, 1, &spill.spill_serial_no)?;
        worksheet.write_string(row, 2, &spill.facility_or_equipment)?;
        worksheet.write_string(row, 3, &spill.location)?;
        worksheet.write_string(row, 4, &spill.incident_date.format("%Y-%m-%d %H:%M:%S").to_string())?;
        if let Some(coords) = &spill.coordinates {
            worksheet.write_string(row, 5, coords)?;
        }
        if let Some(lat) = spill.latitude {
            worksheet.write_number(row, 6, lat)?;
        }
        if let Some(long) = spill.longitude {
            worksheet.write_number(row, 7, long)?;
        }
        worksheet.write_string(row, 8, &spill.spill_status)?;
        worksheet.write_string(row, 9, &spill.description_of_spill_causes)?;
        worksheet.write_number(row, 10, spill.estimated_qty_spilled_or_leaked)?;
        worksheet.write_string(row, 11, &spill.severity)?;
        worksheet.write_string(row, 12, &spill.extent_of_pollution)?;
        worksheet.write_string(row, 13, &spill.precaution_measures)?;
        worksheet.write_string(row, 14, &spill.client_id)?;
        if let Some(oml) = spill.opl_oml_no {
            worksheet.write_number(row, 15, oml as f64)?;
        }
        if let Some(observed) = spill.date_of_incident_or_spill_observed {
            worksheet.write_string(row, 16, &observed.format("%Y-%m-%d %H:%M:%S").to_string())?;
        }
        if let Some(occurred) = spill.date_of_incident_or_spill_occurred {
            worksheet.write_string(row, 17, &occurred.format("%Y-%m-%d %H:%M:%S").to_string())?;
        }
        if let Some(cause) = &spill.incident_cause_of_spill_or_leakage {
            worksheet.write_string(row, 18, cause)?;
        }
        if let Some(category) = &spill.incident_type_of_spill_category {
            worksheet.write_string(row, 19, category)?;
        }
        if let Some(town) = &spill.nearest_town {
            worksheet.write_string(row, 20, town)?;
        }
        if let Some(area) = &spill.operational_area {
            worksheet.write_string(row, 21, area)?;
        }
        if let Some(other_cause) = &spill.other_cause_of_spill_or_leak {
            worksheet.write_string(row, 22, other_cause)?;
        }
        if let Some(other_type) = &spill.other_type_of_spill {
            worksheet.write_string(row, 23, other_type)?;
        }
        if let Some(state) = &spill.state {
            worksheet.write_string(row, 24, state)?;
        }
        if let Some(time_observed) = &spill.time_of_incident_or_spill_observed {
            worksheet.write_string(row, 25, time_observed)?;
        }
        if let Some(time_occurred) = &spill.time_of_incident_or_spill_occurred {
            worksheet.write_string(row, 26, time_occurred)?;
        }
        if let Some(op_type) = &spill.type_of_operation_at_spill_site {
            worksheet.write_string(row, 27, op_type)?;
        }
        if let Some(coord_format) = &spill.coordinate_format {
            worksheet.write_string(row, 28, coord_format)?;
        }
        worksheet.write_string(row, 29, &spill.created_at.format("%Y-%m-%d %H:%M:%S").to_string())?;
        worksheet.write_string(row, 30, &spill.updated_at.format("%Y-%m-%d %H:%M:%S").to_string())?;
    }

    // Auto-fit columns
    for col in 0..headers.len() {
        worksheet.set_column_width(col as u16, 15.0)?;
    }

    workbook.save(filename)?;
    Ok(())
}
