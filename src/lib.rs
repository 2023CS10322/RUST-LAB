#![allow(warnings)]

pub mod parser;
pub mod sheet;

// Export the CLI functions for tests to use
#[cfg(feature = "cli_app")]
pub mod cli_app {
    use crate::parser::*;
    use crate::sheet::*;
    
    // Direct implementation of the functions needed for testing
    pub fn col_to_letters(mut col: i32) -> String {
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

    pub fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
        if *start_row > total_rows {
            *start_row -= 10;
        } else if *start_row > (total_rows - 10) {
            *start_row = total_rows - 10;
        } else if *start_row < 0 {
            *start_row = 0;
        }
    }

    pub fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
        if *start_col > total_cols {
            *start_col -= 10;
        } else if *start_col > (total_cols - 10) {
            *start_col = total_cols - 10;
        } else if *start_col < 0 {
            *start_col = 0;
        }
    }
    
    pub fn process_command(sheet: &mut Box<Spreadsheet>, cmd: &str, status_msg: &mut String) {
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
            sheet.cache.clear();
            sheet.dirty_cells.clear();
            clear_range_cache();
            *status_msg = "Cache cleared".to_string();
        } else if cmd == "undo" {
            #[cfg(feature = "undo_state")]
            {
                sheet.undo(status_msg);
            }
            #[cfg(not(feature = "undo_state"))]
            {
                *status_msg = "Undo feature is not enabled.".to_string();
            }
        } else if cmd == "redo" {
            #[cfg(feature = "undo_state")]
            {
                sheet.redo(status_msg);
            }
            #[cfg(not(feature = "undo_state"))]
            {
                *status_msg = "Undo/Redo feature is not enabled.".to_string();
            }
        } else if cmd.contains('=') {
            if let Some(eq_pos) = cmd.find('=') {
                let cell_name = &cmd[..eq_pos];
                let expr = &cmd[eq_pos + 1..];
                if let Some((row, col)) = cell_name_to_coords(cell_name) {
                    if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                        *status_msg = "Cell out of bounds".to_string();
                    } else {
                        sheet.update_cell_formula(row, col, expr, status_msg);
                    }
                } else {
                    *status_msg = "Invalid cell".to_string();
                }
            }
        } else if cmd.starts_with("history") {
            let parts: Vec<&str> = cmd.split_whitespace().collect();
            if parts.len() == 2 {
                #[cfg(feature = "cell_history")]
                {
                    sheet.skip_default_display = true;
                    *status_msg = "History displayed".to_string();
                }
                #[cfg(not(feature = "cell_history"))]
                {
                    *status_msg = "Cell history feature is not enabled.".to_string();
                }
            }
        } else {
            *status_msg = "unrecognized cmd".to_string();
        }
    }
}

#[cfg(feature = "gui_app")]
pub mod gui_app {
    use crate::parser::*;
    use crate::sheet::*;
    
    // Implement GUI-related functions needed for testing
    pub fn col_to_letters(mut col: i32) -> String {
        if col < 0 {
            return String::new();
        }
        let mut buf = Vec::new();
        loop {
            let remainder = col % 26;
            buf.push((remainder as u8 + b'A') as char);
            col = col / 26 - 1;
            if col < 0 {
                break;
            }
        }
        buf.reverse();
        buf.into_iter().collect()
    }

    pub fn coords_to_cell_name(row: i32, col: i32) -> String {
        let mut n = col + 1;
        let mut col_str = String::new();
        while n > 0 {
            let rem = (n - 1) % 26;
            col_str.push((b'A' + rem as u8) as char);
            n = (n - 1) / 26;
        }
        let col_name: String = col_str.chars().rev().collect();
        format!("{}{}", col_name, row + 1)
    }
}
