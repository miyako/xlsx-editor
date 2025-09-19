use std::env;
use std::error::Error;
use std::fs;
use std::path::Path;
use std::io::{self, Read};

use serde::Deserialize;
use serde_json::Value;
use umya_spreadsheet::{reader, writer};
use umya_spreadsheet::{Style, Cell, HorizontalAlignmentValues, VerticalAlignmentValues, BorderStyleValues};

// JSON spec for a border
#[derive(Debug, Deserialize)]
struct BorderSpec {
    style: Option<String>,
    color: Option<String>,
}

// JSON spec for a single edit
#[derive(Debug, Deserialize)]
struct EditSpec {
    sheet: String,
    cell: String,
    value: Value, // allow string, number, bool, null, or even object/array
#[serde(default)]
    format: Option<String>, // optional number/date/other format
    formula: Option<String>, 
    bold: Option<bool>, 
    italic: Option<bool>, 
    size: Option<f64>,
    font: Option<String>, 
    stroke: Option<String>, 
    fill: Option<String>, 
    halign: Option<String>, 
    valign: Option<String>, 
    left: Option<BorderSpec>, 
    right: Option<BorderSpec>, 
    top: Option<BorderSpec>, 
    bottom: Option<BorderSpec>, 
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
        let cell = sheet.get_cell_mut(cell_ref);
        
        // Set formula if specified
        if let Some(f) = &edit.formula {
            cell.set_formula(f);
        }//else{
        // Set value based on type
        match &edit.value {
            Value::String(s) => {
                cell.set_value(s);
            }
            Value::Number(num) => {
                // write number as string; umya will detect numeric strings and store as number. 
                cell.set_value(num.to_string());
            }
            Value::Bool(b) => {
                cell.set_value_bool(*b);
            }
            Value::Null => {
                cell.set_value("");
            }
            other => {
                let as_text = serde_json::to_string(other)?;
                cell.set_value(as_text);
            }
        }            
        // }
                
        apply_font(cell, &edit);
        
        // Apply optional format if provided
        if let Some(fmt) = &edit.format {
            eprintln!("set format code of {} to '{}'", edit.cell, fmt);
            let mut style = Style::default();
            style.get_number_format_mut().set_format_code(fmt);
            cell.set_style(style);
        }        
        
    }

    // Save to output file
    let out_path = Path::new(output_file);
    writer::xlsx::write(&book, out_path)?;

    println!("Wrote {} edits -> {}", edits.len(), output_file);
    Ok(())
}

fn apply_font(cell: &mut Cell, edit: &EditSpec) {
    
    let mut style = cell.get_style().clone();
            
    let font = style.get_font_mut();
    
    if let Some(b) = edit.bold {
        eprintln!("set bold of {} to {}", edit.cell, b);
        font.set_bold(b);
    }
    
    if let Some(i) = edit.italic {
        eprintln!("set italic of {} to {}", edit.cell, i);
        font.set_italic(i);
    }
    
    if let Some(s) = edit.size {
        eprintln!("set size of {} to {}", edit.cell, s);
        font.set_size(s);
    }
    
    if let Some(n) = &edit.font {
        eprintln!("set font of {} to {}", edit.cell, n);
        font.set_name(n);
    }
    
    if let Some(c) = &edit.stroke {
        eprintln!("set stroke of {} to {}", edit.cell, c);
        font.get_color_mut().set_argb(c);
    }
    
    if let Some(f) = &edit.fill {
        eprintln!("set fill of {} to {}", edit.cell, f);
        let pattern_fill = style.get_fill_mut().get_pattern_fill_mut();
        pattern_fill.set_pattern_type(umya_spreadsheet::PatternValues::Solid); 
        pattern_fill.get_foreground_color_mut().set_argb(f);  
    }
    
    let borders = style.get_borders_mut();
    
    if let Some(l) = &edit.left {
        let left = borders.get_left_mut();
        if let Some(s) = &l.style {
            eprintln!("set left border of {} to {}", edit.cell, s);
            left.set_style(parse_border_style(s));
        }
        if let Some(c) = &l.color {
            eprintln!("set left border color of {} to {}", edit.cell, c);
            left.get_color_mut().set_argb(c);        
        }
    }
    
    if let Some(r) = &edit.right {
        let right = borders.get_right_mut();
        if let Some(s) = &r.style {
            println!("set right border of {} to {}", edit.cell, s);
            right.set_style(parse_border_style(s));
        }   
        if let Some(c) = &r.color {
            eprintln!("set right border color of {} to {}", edit.cell, c);
            right.get_color_mut().set_argb(c);        
        }
    }
    
    if let Some(t) = &edit.top {
        let top = borders.get_top_mut();
        if let Some(s) = &t.style {
            println!("set top border of {} to {}", edit.cell, s);
            top.set_style(parse_border_style(s));
        }
        if let Some(c) = &t.color {
            eprintln!("set top border color of {} to {}", edit.cell, c);
            top.get_color_mut().set_argb(c);        
        }
    }
        
    if let Some(b) = &edit.bottom {
        let bottom = borders.get_bottom_mut();
        if let Some(s) = &b.style {
            println!("set bottom border of {} to {}", edit.cell, s);
            bottom.set_style(parse_border_style(s));
        }
        if let Some(c) = &b.color {
            eprintln!("set bottom border color of {} to {}", edit.cell, c);
            bottom.get_color_mut().set_argb(c);        
        }
    }
    
    let alignment = style.get_alignment_mut();
    
    // Horizontal alignment
    if let Some(halign) = &edit.halign {
        let h_enum = match halign.to_lowercase().as_str() {
            "left" => HorizontalAlignmentValues::Left,
            "center" => HorizontalAlignmentValues::Center,
            "right" => HorizontalAlignmentValues::Right,
            "fill" => HorizontalAlignmentValues::Fill,
            "justify" => HorizontalAlignmentValues::Justify,
            "continuous" => HorizontalAlignmentValues::CenterContinuous,
            "distributed" => HorizontalAlignmentValues::Distributed,
            other => {
                eprintln!("Unknown alignment value '{}', defaulting to Left", other);
                HorizontalAlignmentValues::Left
            }
        };
        eprintln!("set horizontal alignment of {} to {}", edit.cell, halign);
        alignment.set_horizontal(h_enum);
    }
    
    // Vertical alignment
    if let Some(valign) = &edit.valign {
        let v_enum = match valign.to_lowercase().as_str() {
            "top" => VerticalAlignmentValues::Top,
            "center" => VerticalAlignmentValues::Center,
            "bottom" => VerticalAlignmentValues::Bottom,
            "justify" => VerticalAlignmentValues::Justify,
            "distributed" => VerticalAlignmentValues::Distributed,
            other => {
                eprintln!("Unknown alignment value '{}', defaulting to Bottom", other);
                VerticalAlignmentValues::Bottom
            }
        };
        eprintln!("set vertical alignment of {} to {}", edit.cell, valign);
        alignment.set_vertical(v_enum);
    }
   
    cell.set_style(style);
}

fn parse_border_style(s: &str) -> BorderStyleValues {
    
    match s.to_lowercase().as_str() {
        "none" => BorderStyleValues::None,
        "thin" => BorderStyleValues::Thin,
        "medium" => BorderStyleValues::Medium,
        "dashed" => BorderStyleValues::Dashed,
        "dotted" => BorderStyleValues::Dotted,
        "thick" => BorderStyleValues::Thick,
        "double" => BorderStyleValues::Double,
        "hair" => BorderStyleValues::Hair,
        "mediumdashed" => BorderStyleValues::MediumDashed,
        "dashdot" => BorderStyleValues::DashDot,
        "mediumdashdot" => BorderStyleValues::MediumDashDot,
        "dashdotdot" => BorderStyleValues::DashDotDot,
        "mediumdashdotdot" => BorderStyleValues::MediumDashDotDot,
        "slantdashdot" => BorderStyleValues::SlantDashDot,
        other => {
            eprintln!("Unknown border style '{}', defaulting to None", other);
            BorderStyleValues::None
        }
    }
}