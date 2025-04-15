#![allow(warnings)]
mod parser;
mod sheet;

use parser::*;
use sheet::*;
use std::env;
use std::io::{self, Write};
use std::time::Instant;
use std::time::Duration;

// Global variable as in C (check)
static mut CHECK: bool = false;

// Converts a 0-indexed column number into its corresponding letter string.
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

// Clamps vertical viewport.
fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
    if *start_row > total_rows {
        *start_row -= 10;
    } else if *start_row > (total_rows - 10) {
        *start_row = total_rows - 10;
    } else if *start_row < 0 {
        *start_row = 0;
    }
}

// Clamps horizontal viewport.
fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
    if *start_col > total_cols {
        *start_col -= 10;
    } else if *start_col > (total_cols - 10) {
        *start_col = total_cols - 10;
    } else if *start_col < 0 {
        *start_col = 0;
    }
}

// Displays the grid (viewport 10x10).
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
            let status = sheet.get_cell_status(r, c);
            if status == CellStatus::Error {
                print!("{:<12}", "ERR");
            } else {
                print!("{:<12}", sheet.get_cell_value(r, c));
            }
        }
        println!();
    }
}

// Displays grid from a specified start.
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
            let status = sheet.get_cell_status(r, c);
            if status == CellStatus::Error {
                print!("{:<12}", "ERR");
            } else {
                print!("{:<12}", sheet.get_cell_value(r, c));
            }
        }
        println!();
    }
}

// Process commands: scrolling, cell assignment, output control.
fn process_command(sheet: &mut Spreadsheet, cmd: &str, status_msg: &mut String) {
    if cmd == "w" {
        sheet.top_row -= 10;
        clamp_viewport_ve(sheet.total_rows, &mut sheet.top_row);
    } else if cmd == "s" {
        sheet.top_row += 10;
        clamp_viewport_ve(sheet.total_rows, &mut sheet.top_row);
    } else if cmd == "a" {
        sheet.left_col -= 10;
        clamp_viewport_hz(sheet.total_cols, &mut sheet.left_col);
    } else if cmd == "d" {
        sheet.left_col += 10;
        clamp_viewport_hz(sheet.total_cols, &mut sheet.left_col);
    } else if cmd.starts_with("scroll_to") {
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        if parts.len() == 2 {
            let cell_name = parts[1];
            if let Some((row, col)) = cell_name_to_coords(cell_name) {
                if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                    *status_msg = "Cell reference out of bounds".to_string();
                } else {
                    sheet.top_row = row;
                    sheet.left_col = col;
                }
            } else {
                *status_msg = "Invalid cell".to_string();
            }
        } else {
            *status_msg = "Invalid command".to_string();
        }
    } else if cmd == "disable_output" {
        sheet.output_enabled = false;
    } else if cmd == "enable_output" {
        sheet.output_enabled = true;
    } else if cmd == "clear_cache" {
        // Clear both sheet cache and parser cache
        sheet.cache.clear();
        sheet.dirty_cells.clear();
        parser::clear_range_cache();
        *status_msg = "Cache cleared".to_string();
    } else if cmd.contains('=') {
        if let Some(eq_pos) = cmd.find('=') {
            let cell_name = &cmd[..eq_pos];
            let expr = &cmd[eq_pos + 1..];
            if let Some((row, col)) = cell_name_to_coords(cell_name) {
                if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                    *status_msg = "Cell out of bounds".to_string();
                } else {
                    // Call update_cell_formula.
                    sheet.update_cell_formula(row, col, expr, status_msg);
                }
            } else {
                *status_msg = "Invalid cell".to_string();
            }
        }
    } else {
        *status_msg = "unrecognized cmd".to_string();
    }
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: {} <rows> <cols>", args[0]);
        return;
    }
    let rows: i32 = args[1].parse().unwrap_or(0);
    let cols: i32 = args[2].parse().unwrap_or(0);
    if rows < 1 || cols < 1 {
        eprintln!("Invalid dimensions.");
        return;
    }
    let mut cmd = String::new();
    let mut status_msg = String::from("ok");
    let mut elapsed_time = 0.0;

    // Allocate the spreadsheet on the heap.
    let mut sheet = Box::new(Spreadsheet::new(rows, cols));
    println!("Boxed sheet at address {:p}, rows={}, cols={}", &*sheet, sheet.total_rows, sheet.total_cols);


    display_grid(&sheet);
    print!("[{:.1}] ({}) > ", elapsed_time, status_msg);
    io::stdout().flush().unwrap();

    loop {
        cmd.clear();
        if io::stdin().read_line(&mut cmd).is_err() {
            status_msg = "Invalid command".to_string();
        }
        let cmd = cmd.trim();
        if cmd == "q" {
            break;
        }
        let start = Instant::now();
        // Pass a mutable reference to the spreadsheet.
        process_command(&mut *sheet, cmd, &mut status_msg);
        let duration = start.elapsed();
        elapsed_time = duration.as_secs_f64();

        if sheet.output_enabled && cmd != "enable_output" {
            display_grid_from(&sheet, sheet.top_row, sheet.left_col);
        }
        
        print!("[{:.1}] ({}) > ", elapsed_time, status_msg);
        io::stdout().flush().unwrap();
        status_msg = "ok".to_string();
    }
}

