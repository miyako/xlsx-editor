use std::env;
use std::error::Error;
use std::fs;
use std::path::Path;
use std::io::{self, Read};

use serde::Deserialize;
use serde_json::Value;
use umya_spreadsheet::{reader, writer};

// JSON spec for a single edit
#[derive(Debug, Deserialize)]
struct EditSpec {
    sheet: String,
    cell: String,
    value: Value, // allow string, number, bool, null, or even object/array
}

fn main() -> Result<(), Box<dyn Error>> {

    let args: Vec<String> = env::args().collect();
    
    if args.len() < 3 || args.len() > 4 {
        eprintln!(
            "Usage: {} <input.xlsx|-> <output.xlsx> [spec.json]",
            args[0]
        );
        eprintln!("Use - as input to start from a new blank workbook.");
        eprintln!("If spec.json is omitted, JSON will be read from stdin.");
        std::process::exit(1);
    }

    // Read spec JSON (from file or stdin)
    let raw = if args.len() == 4 {
        let spec_file = &args[3];
        fs::read_to_string(spec_file).expect("Failed to read spec.json")
    } else {
        let mut buffer = String::new();
        io::stdin()
            .read_to_string(&mut buffer)
            .expect("Failed to read from stdin");
        buffer
    };

    let input_file = &args[1]; 
    let output_file = &args[2];

    // Try to parse as an array first; otherwise, parse a single object
    let mut edits: Vec<EditSpec> = match serde_json::from_str(&raw) {
        Ok(v) => v,
        Err(_) => {
            // try single object
            let single: EditSpec = serde_json::from_str(&raw)?;
            vec![single]
        }
    };

    // Load or create workbook
    let mut book = if input_file == "-" {
        umya_spreadsheet::new_file()
    } else {
        let p = Path::new(input_file);
        if p.exists() {
            reader::xlsx::read(p)?
        } else {
            // create new file if input not found
            umya_spreadsheet::new_file()
        }
    };

    // Apply each edit
    for edit in edits.iter_mut() {
        // ensure the sheet exists; create if missing
        if book.get_sheet_by_name_mut(&edit.sheet).is_none() {
            // new_sheet returns something but we just need the sheet to exist
            let _ = book.new_sheet(&edit.sheet);
        }

        // safe to unwrap now
        let sheet = book.get_sheet_by_name_mut(&edit.sheet).expect("sheet exists");

        let cell_ref: &str = edit.cell.as_str();

        match &edit.value {
            Value::String(s) => {
                sheet.get_cell_mut(cell_ref).set_value(s);
            }
            Value::Number(num) => {
                // write number as string; umya will detect numeric strings and store as number.
                sheet.get_cell_mut(cell_ref).set_value(num.to_string());
            }
            Value::Bool(b) => {
                sheet.get_cell_mut(cell_ref).set_value_bool(*b);
            }
            Value::Null => {
                sheet.get_cell_mut(cell_ref).set_value("");
            }
            other => {
                // For arrays/objects: store the JSON string representation
                let as_text = serde_json::to_string(other)?;
                sheet.get_cell_mut(cell_ref).set_value(as_text);
            }
        }
    }

    // Save to output file
    let out_path = Path::new(output_file);
    writer::xlsx::write(&book, out_path)?;

    println!("Wrote {} edits -> {}", edits.len(), output_file);
    Ok(())
}
