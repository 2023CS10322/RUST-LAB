#![allow(warnings)] // Keep if needed
mod parser;
mod sheet;

use crate::sheet::cell_name_to_coords; // Ensure this function is accessible [1, 3]
use parser::*; // Assuming parser has necessary items [2]
use sheet::*; // Assuming sheet has Spreadsheet, CellStatus, get_cell_history [1]
use std::env;
use std::io::{self, Write};
use std::time::Instant;
// Remove: use std::time::Duration; - Not directly used in this modification

// Global variable as in C (check) - This is generally discouraged in Rust
// static mut CHECK: bool = false;

// Converts a 0-indexed column number into its corresponding letter string. [3]
fn col_to_letters(mut col: i32) -> String {
    let mut buf = Vec::new();
    loop {
        buf.push(((col % 26) as u8 + b'A') as char);
        col = col / 26 - 1;
        if col < 0 {
            break;
        }
    }
    buf.reverse();
    buf.into_iter().collect()
}

// Clamps vertical viewport. [3]
fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
    if *start_row > total_rows {
        *start_row -= 10;
    } else if *start_row > (total_rows - 10) {
        *start_row = total_rows - 10;
    } else if *start_row < 0 {
        *start_row = 0;
    }
}

// Clamps horizontal viewport. [3]
fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
    if *start_col > total_cols {
        *start_col -= 10;
    } else if *start_col > (total_cols - 10) {
        *start_col = total_cols - 10;
    } else if *start_col < 0 {
        *start_col = 0;
    }
}

// Displays the grid (viewport 10x10). [3]
// Note: Does not use sheet.output_enabled, display_grid_from does
fn display_grid(sheet: &Spreadsheet) {
    let start_row = sheet.top_row;
    let start_col = sheet.left_col;
    let mut end_row = start_row + 10;
    let mut end_col = start_col + 10;
    if end_row > sheet.total_rows {
        end_row = sheet.total_rows;
    }
    if end_col > sheet.total_cols {
        end_col = sheet.total_cols;
    }

    // Print column headers.
    print!("     ");
    for c in start_col..end_col {
        let col_buf = col_to_letters(c);
        print!("{:<12}", col_buf);
    }
    println!();

    // Print rows with values.
    for r in start_row..end_row {
        print!("{:<4} ", r + 1);
        for c in start_col..end_col {
            // Get cell value from the sparse representation
            let status = sheet.get_cell_status(r, c); // [1]
            if status == CellStatus::Error { // [1]
                print!("{:<12}", "ERR");
            } else {
                print!("{:<12}", sheet.get_cell_value(r, c)); // [1]
            }
        }
        println!();
    }
}

// Displays grid from a specified start. [3]
fn display_grid_from(sheet: &Spreadsheet, start_row: i32, start_col: i32) {
    // Calculate max displayable rows/columns
    let mut max_col = start_col + 10;
    if max_col > sheet.total_cols {
        max_col = sheet.total_cols;
    }

    let mut max_row = start_row + 10;
    if max_row > sheet.total_rows {
        max_row = sheet.total_rows;
    }

    // Always print at least column headers
    print!("     ");
    for c in start_col..max_col {
        let col_buf = col_to_letters(c);
        print!("{:<12}", col_buf);
    }
    println!();

    // Print rows with boundary checking
    for r in start_row..max_row {
        if r < 0 || r >= sheet.total_rows {
            continue;
        }

        print!("{:<4} ", r + 1);
        for c in start_col..max_col {
            if c < 0 || c >= sheet.total_cols {
                print!("{:<12}", "--");
                continue;
            }

            // Get cell value from the sparse representation
            let status = sheet.get_cell_status(r, c); // [1]
            if status == CellStatus::Error { // [1]
                print!("{:<12}", "ERR");
            } else {
                print!("{:<12}", sheet.get_cell_value(r, c)); // [1]
            }
        }
        println!();
    }
}

// Process commands: scrolling, cell assignment, output control. [3]
fn process_command(sheet: &mut Spreadsheet, cmd: &str, status_msg: &mut String) {
    status_msg.clear(); // Clear previous status at the start
    let cmd_lower = cmd.to_lowercase(); // Use lowercase for command matching

    if cmd_lower == "w" {
        sheet.top_row -= 10;
        clamp_viewport_ve(sheet.total_rows, &mut sheet.top_row);
        status_msg.push_str("ok");
    } else if cmd_lower == "s" {
        sheet.top_row += 10;
        clamp_viewport_ve(sheet.total_rows, &mut sheet.top_row);
        status_msg.push_str("ok");
    } else if cmd_lower == "a" {
        sheet.left_col -= 10;
        clamp_viewport_hz(sheet.total_cols, &mut sheet.left_col);
        status_msg.push_str("ok");
    } else if cmd_lower == "d" {
        sheet.left_col += 10;
        clamp_viewport_hz(sheet.total_cols, &mut sheet.left_col);
        status_msg.push_str("ok");
    } else if cmd_lower.starts_with("scroll_to") {
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        if parts.len() == 2 {
            let cell_name = parts[1];
            if let Some((row, col)) = cell_name_to_coords(cell_name) { // [1, 3]
                if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                    status_msg.push_str("Cell reference out of bounds");
                } else {
                    sheet.top_row = row;
                    sheet.left_col = col;
                    status_msg.push_str("ok");
                }
            } else {
                status_msg.push_str("Invalid cell");
            }
        } else {
            status_msg.push_str("Usage: scroll_to <CellReference>");
        }
    } else if cmd_lower == "disable_output" {
        sheet.output_enabled = false;
        status_msg.push_str("Output disabled");
    } else if cmd_lower == "enable_output" {
        sheet.output_enabled = true;
        status_msg.push_str("Output enabled");
        // Need to display grid here if output was previously off
        sheet.skip_default_display = false; // Ensure next loop iteration displays
    } else if cmd_lower == "clear_cache" {
        sheet.cache.clear(); // [1]
        sheet.dirty_cells.clear(); // [1]
        parser::clear_range_cache(); // [2]
        status_msg.push_str("Cache cleared");

    // --- Add history command handling ---
    } else if cmd_lower.starts_with("history") {
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        if parts.len() == 2 {
            let cell_ref = parts[1];
            if let Some((row, col)) = cell_name_to_coords(cell_ref) { // [1, 3]
                if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                     status_msg.push_str(&format!("Cell {} out of bounds", cell_ref.to_uppercase()));
                } else {
                    // --- Feature Check ---
                    #[cfg(feature = "cell_history")] // [4]
                    {
                        // Assuming get_cell_history exists in Spreadsheet impl [1]
                        match sheet.get_cell_history(row, col) {
                            Some(history) if !history.is_empty() => {
                                // Print directly instead of using status_msg due to potential length
                                println!("History for {}:", cell_ref.to_uppercase());
                                // Print oldest first (index 1)
                                for (i, val) in history.iter().enumerate() {
                                    println!("  {}: {}", i + 1, val);
                                }
                                let current_val = sheet.get_cell_value(row, col);
                                println!("  Current: {}", current_val);
                                status_msg.push_str("History displayed"); // Set status message
                            }
                            _ => { // Cell exists but has no history, or cell doesn't exist yet
                                status_msg.push_str(&format!("No recorded history for {}", cell_ref.to_uppercase()));
                            }
                        }
                         sheet.skip_default_display = true; // Don't redisplay grid after history
                    }
                    #[cfg(not(feature = "cell_history"))] // [4]
                    {
                         status_msg.push_str("Cell history feature is not enabled.");
                         // sheet.skip_default_display = true; // Prevent grid redraw
                    }
                    // --- End Feature Check ---
                }
            } else {
                status_msg.push_str(&format!("Invalid cell reference: {}", cell_ref));
            }
        } else {
            status_msg.push_str("Usage: history <CellReference>");
        }
    // --- End history command handling ---

    // --- Add undo/redo command handling ---
     } else if cmd_lower == "undo" {
          // --- Feature Check ---
         #[cfg(feature = "undo_state")] // [6, 8, 9]
         {
             sheet.undo(status_msg); // Call the undo method [1]
             // status_msg is set within the undo method
         }
         #[cfg(not(feature = "undo_state"))] // [6, 8, 9]
         {
             status_msg.push_str("Undo feature is not enabled.");
         }
         // --- End Feature Check ---
     } else if cmd_lower == "redo" {
          // --- Feature Check ---
          #[cfg(feature = "undo_state")] // <-- Update feature name [1, 3]
          {
              sheet.redo(status_msg); // Call the redo method (sets status_msg) [1]
          }
          #[cfg(not(feature = "undo_state"))] // <-- Update feature name [1, 3]
          {
              status_msg.push_str("Undo/Redo feature is not enabled.");
          }
          // --- End Feature Check ---
     // --- End undo/redo command handling ---


    } else if cmd.contains('=') {
        // Make case-insensitive for cell name part? Assume original case sensitivity for now.
        if let Some(eq_pos) = cmd.find('=') {
            let cell_name = cmd[..eq_pos].trim(); // Trim spaces around cell name
            let expr = cmd[eq_pos + 1..].trim(); // Trim spaces around expression
            if let Some((row, col)) = cell_name_to_coords(cell_name) { // [1, 3]
                if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                    status_msg.push_str("Cell out of bounds");
                } else {
                    // Call update_cell_formula. It sets status_msg internally. [1]
                    sheet.update_cell_formula(row, col, expr, status_msg);
                }
            } else {
                status_msg.push_str("Invalid cell reference");
            }
        } else {
             // Should not happen if '=' is present, but handle defensively
             status_msg.push_str("Invalid assignment format");
        }
    } else if !cmd.is_empty() { // Only flag non-empty unrecognized commands
        status_msg.push_str("Unrecognized command");
    } else {
        // Empty input, just redisplay prompt with "ok"
        status_msg.push_str("ok");
    }

    // If status_msg wasn't set by any branch, default to "ok" (unless empty input)
    if status_msg.is_empty() && !cmd.is_empty() {
        status_msg.push_str("ok");
    }
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: {} <rows> <cols>", args[0]);
        return;
    }
    // Use expect for more direct error handling on parse failure
    let rows: i32 = args[1].parse().expect("Invalid number of rows");
    let cols: i32 = args[2].parse().expect("Invalid number of columns");

    if rows < 1 || cols < 1 {
        eprintln!("Rows and columns must be positive numbers.");
        return;
    }

    let mut cmd = String::new();
    // Initialize status_msg for the first prompt
    let mut status_msg = String::from("Ready");
    let mut elapsed_time; // No need to initialize here

    // Allocate the spreadsheet on the heap. [3]
    let mut sheet = Spreadsheet::new(rows, cols); // [1]
    println!(
        "Spreadsheet created: {} rows, {} columns",
        sheet.total_rows, sheet.total_cols
    );


    // Initial display
    if sheet.output_enabled {
        display_grid(&sheet);
    }
    print!("[0.0] ({}) > ", status_msg); // Show initial status
    io::stdout().flush().unwrap();

    loop {
        cmd.clear();
        match io::stdin().read_line(&mut cmd) {
            Ok(0) => break, // EOF reached
            Ok(_) => { /* Line read successfully */ }
            Err(e) => {
                eprintln!("Error reading input: {}", e);
                break; // Exit on input error
            }
        }

        let cmd_trimmed = cmd.trim(); // Use trimmed version for processing
        if cmd_trimmed.eq_ignore_ascii_case("q") { // Case-insensitive quit
            break;
        }

        // Reset status message for the new command processing
        status_msg.clear();
        sheet.skip_default_display = false; // Reset display skip flag

        let start = Instant::now();
        // Pass a mutable reference to the spreadsheet. [3]
        process_command(&mut sheet, cmd_trimmed, &mut status_msg);
        let duration = start.elapsed();
        // Display time in milliseconds for potentially faster operations
        elapsed_time = duration.as_secs_f64() * 1000.0;

        // Display grid unless output is off or command suppressed it
        if sheet.output_enabled && !sheet.skip_default_display {
            // Use display_grid_from for consistent viewport handling
            display_grid_from(&sheet, sheet.top_row, sheet.left_col);
        }

        // Print prompt
        print!("[{:.1}] ({}) > ", elapsed_time, status_msg);
        io::stdout().flush().unwrap();
    }

    println!("\nExiting.");
}
