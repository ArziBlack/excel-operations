use dotenv::dotenv;
use std::env;
use xcel_operations::spill_service::{fetch_spills, save_spills_to_excel, ApiError};

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load environment variables from .env file
    dotenv().ok();
    let client_id = env::var("F3D_CLIENT_ID").expect("F3D_CLIENT_ID environment variable not set");
    let token = env::var("F3D_TOKEN").expect("F3D_TOKEN environment variable not set");

    println!("Fetching spill data...");
    match fetch_spills(&client_id, &token).await.map_err(|e| Box::new(e) as Box<dyn std::error::Error>) {
        Ok(response) => {
            if response.success {
                println!("Successfully fetched {} spills", response.data.len());
                if !response.data.is_empty() {
                    match save_spills_to_excel(&response.data, "spill_register.xlsx") {
                        Ok(_) => println!("Successfully saved spills to Excel file"),
                        Err(e) => {
                            eprintln!("Error saving spills to Excel: {}", e);
                            return Err(e);
                        }
                    }
                } else {
                    println!("No spills found");
                }
            } else {
                println!("Failed to fetch spill data: {}", response.message);
            }
        },
        Err(e) => {
            let error_msg = match e.downcast_ref::<ApiError>() {
                Some(api_error) => match api_error {
                    ApiError::RequestFailed(error_response) => {
                        if let Some(error) = &error_response.error {
                            format!("API request failed:\n  Message: {}\n  Error details: {}",
                                error_response.message, error)
                        } else {
                            format!("API request failed: {}", error_response.message)
                        }
                    },
                    ApiError::NetworkError(e) => format!("Network error: {}", e),
                    ApiError::Other(e) => format!("Other error: {}", e),
                },
                None => format!("Unknown error: {}", e),
            };
            println!("{}", error_msg);
            return Err(e);
        }
    }
    
    Ok(())
}
