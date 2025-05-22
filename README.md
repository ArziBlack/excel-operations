# Excel Operations with Rust

A Rust application for managing spill incident data and Excel operations. This project provides functionality to fetch spill data from an API, save it to Excel files, and manage spill incident records.

## Features

- Fetch spill incident data from API
- Save spill data to Excel files
- Create and manage spill incident records
- Read, write, and modify Excel files
- Handle various spill incident attributes and metadata
- **Dynamic Data Processing**: Dynamically read headers from Excel files and extract data matching specific struct fields (e.g., `SpillIncident` struct).

## Prerequisites

- Rust (latest stable version)
- Cargo (comes with Rust)
- Environment variables:
  - `F3D_CLIENT_ID`: Your client ID for the API
  - `F3D_TOKEN`: Your authentication token for the API

## Installation

1. Clone the repository:
```bash
git clone https://github.com/ArziBlack/excel-operations.git
cd excel-operations
```

2. Create a `.env` file in the project root and add your credentials:
```env
F3D_CLIENT_ID=your_client_id
F3D_TOKEN=your_token
```

3. Build the project:
```bash
cargo build
```

## Usage

### Fetch Spill Data
To fetch spill data from the API and save it to Excel:
```bash
cargo run --bin fetch_spills
```

### Create Example Spill Records
To run the spill example that creates sample spill records:
```bash
cargo run --bin spill_example
```

### Run Excel Operations Examples
To run the main Excel operations examples:
```bash
cargo run
```

### Dynamic Data Processing Example
To run the dynamic data processing example:
```bash
cargo run --bin dynamic_data_processing
```

### Running the Dynamic Spill Loader

To run the dynamic spill loader with default field names:
```bash
cargo run --bin dynamic_spill_loader
```

To specify custom field names to match with the struct:
```bash
cargo run --bin dynamic_spill_loader "Field1" "Field2" "Field3"
```

This will process the Excel file and create `loaded_register_dynamic.xlsx` with the extracted data. The loader now trims whitespace from Excel headers to improve matching accuracy with expected field names.

## Project Structure

- `src/bin/fetch_spills.rs`: Fetches spill data from API
- `src/bin/spill_example.rs`: Example of creating spill incidents
- `src/bin/dynamic_spill_loader.rs`: Binary for dynamic loading of spill data with customizable field matching.
- `src/spill_register.rs`: Core spill incident data structures and functions
- `src/spill_service.rs`: API service and Excel operations for spills
- `src/main.rs`: General Excel operations examples
- `src/dynamic_data_processing.rs`: Dynamic data processing logic

## File Outputs

- `spill_register.xlsx`: Contains spill incident records
- `data.xlsx`: Sample data file for Excel operations examples
- `output.xlsx`: Output file for modified Excel data
- `loaded_register_dynamic.xlsx`: Result of dynamically processing spill data

## Error Handling

The application includes comprehensive error handling for:
- API request failures
- File I/O operations
- Data validation
- Excel operations

## Dependencies

- `reqwest`: HTTP client for API requests
- `rust_xlsxwriter`: Excel file creation and writing
- `calamine`: Excel file reading
- `chrono`: Date and time handling
- `serde`: Data serialization/deserialization
- `dotenv`: Environment variable management

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -am 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
