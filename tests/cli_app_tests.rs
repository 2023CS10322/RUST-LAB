#[cfg(feature = "cli_app")]
#[cfg(test)]
mod cli_tests {
    use spreadsheet::parser::*;
    use spreadsheet::sheet::*;
    use std::collections::HashSet;
    // use std::io::{self, Write};
    // use std::time::Instant;

    // Define the cli_app functions we need for testing
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

    fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
        if *start_row > total_rows {
            *start_row -= 10;
        } else if *start_row > (total_rows - 10) {
            *start_row = total_rows - 10;
        } else if *start_row < 0 {
            *start_row = 0;
        }
    }

    fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
        if *start_col > total_cols {
            *start_col -= 10;
        } else if *start_col > (total_cols - 10) {
            *start_col = total_cols - 10;
        } else if *start_col < 0 {
            *start_col = 0;
        }
    }

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
        } else {
            *status_msg = "unrecognized cmd".to_string();
        }
    }

    // Test helper functions from main.rs
    #[test]
    fn test_col_to_letters() {
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(1), "B");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
        assert_eq!(col_to_letters(51), "AZ");
        assert_eq!(col_to_letters(52), "BA");
        assert_eq!(col_to_letters(701), "ZZ");
        assert_eq!(col_to_letters(702), "AAA");
    }

    #[test]
    fn test_clamp_viewport_functions() {
        // Test vertical clamping
        {
            let mut start_row = -5;
            clamp_viewport_ve(100, &mut start_row);
            assert_eq!(start_row, 0);
            
            start_row = 95;
            clamp_viewport_ve(100, &mut start_row);
            assert_eq!(start_row, 90);
            
            start_row = 150;
            clamp_viewport_ve(100, &mut start_row);
            assert_eq!(start_row, 140);
        }
        
        // Test horizontal clamping
        {
            let mut start_col = -3;
            clamp_viewport_hz(100, &mut start_col);
            assert_eq!(start_col, 0);
            
            start_col = 95;
            clamp_viewport_hz(100, &mut start_col);
            assert_eq!(start_col, 90);
            
            start_col = 120;
            clamp_viewport_hz(100, &mut start_col);
            assert_eq!(start_col, 110);
        }
    }

    // Test process_command function
    #[test]
    fn test_process_command_navigation() {
        let mut sheet = Spreadsheet::new(100, 100);
        let mut status_msg = String::new();
        
        // Test navigation commands
        sheet.top_row = 20;
        sheet.left_col = 20;
        
        // Test 'w' command (up)
        process_command(&mut sheet, "w", &mut status_msg);
        assert_eq!(sheet.top_row, 10);
        
        // Test 's' command (down)
        process_command(&mut sheet, "s", &mut status_msg);
        assert_eq!(sheet.top_row, 20);
        
        // Test 'a' command (left)
        process_command(&mut sheet, "a", &mut status_msg);
        assert_eq!(sheet.left_col, 10);
        
        // Test 'd' command (right)
        process_command(&mut sheet, "d", &mut status_msg);
        assert_eq!(sheet.left_col, 20);
    }

    #[test]
    fn test_process_command_scroll_to() {
        let mut sheet = Spreadsheet::new(100, 100);
        let mut status_msg = String::new();
        
        // Test scroll_to with valid cell
        process_command(&mut sheet, "scroll_to C15", &mut status_msg);
        assert_eq!(sheet.top_row, 14);
        assert_eq!(sheet.left_col, 2);
        
        // Test scroll_to with invalid cell name
        process_command(&mut sheet, "scroll_to X123X", &mut status_msg);
        assert_eq!(status_msg, "Invalid cell");
        
        // Test scroll_to with out of bounds
        process_command(&mut sheet, "scroll_to Z999", &mut status_msg);
        assert_eq!(status_msg, "Cell reference out of bounds");
    }

    #[test]
    fn test_process_command_output_control() {
        let mut sheet = Spreadsheet::new(10, 10);
        let mut status_msg = String::new();
        
        // Test enable/disable output
        assert_eq!(sheet.output_enabled, true);
        
        process_command(&mut sheet, "disable_output", &mut status_msg);
        assert_eq!(sheet.output_enabled, false);
        
        process_command(&mut sheet, "enable_output", &mut status_msg);
        assert_eq!(sheet.output_enabled, true);
    }

    #[test]
    fn test_process_command_clear_cache() {
        let mut sheet = Spreadsheet::new(10, 10);
        let mut status_msg = String::new();
        
        // Add some data to the cache
        sheet.cache.insert("test".to_string(), CachedRange {
            value: 42,
            dependencies: HashSet::new(),
        });
        
        // Clear cache
        process_command(&mut sheet, "clear_cache", &mut status_msg);
        
        // Verify cache is cleared
        assert!(sheet.cache.is_empty());
        assert_eq!(status_msg, "Cache cleared");
    }

    #[test]
    fn test_process_command_cell_assignment() {
        let mut sheet = Spreadsheet::new(10, 10);
        let mut status_msg = String::new();
        
        // Assign value to cell
        process_command(&mut sheet, "A1=42", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        
        // Assign formula to cell
        process_command(&mut sheet, "B1=A1*2", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 1), 84);
        
        // Invalid cell reference
        process_command(&mut sheet, "Z99=10", &mut status_msg);
        assert_eq!(status_msg, "Cell out of bounds");
        
        // Invalid cell name
        process_command(&mut sheet, "ZZZ=10", &mut status_msg);
        assert_eq!(status_msg, "Invalid cell");
    }

    #[cfg(feature = "undo_state")]
    #[test]
    fn test_process_command_undo_redo() {
        let mut sheet = Spreadsheet::new(10, 10);
        let mut status_msg = String::new();
        
        // Make changes
        process_command(&mut sheet, "A1=42", &mut status_msg);
        process_command(&mut sheet, "A1=100", &mut status_msg);
        
        // Test undo
        process_command(&mut sheet, "undo", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        assert!(status_msg.contains("Undo successful"));
        
        // Test redo
        process_command(&mut sheet, "redo", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 100);
        assert!(status_msg.contains("Redo successful"));
    }
}
