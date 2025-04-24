#[cfg(test)]
#[cfg(feature = "cli_app")]
mod tests {
    use spreadsheet::cli_app::{col_to_letters, clamp_viewport_ve, clamp_viewport_hz, process_command};
    use spreadsheet::sheet::*;
    use std::collections::HashSet;
    
    // No need to reimplement functions - using the actual ones from main.rs
    
    // Test the CLI app functionality
    #[test]
    fn test_cli_col_to_letters() {
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
        assert_eq!(col_to_letters(51), "AZ");
        assert_eq!(col_to_letters(52), "BA");
        assert_eq!(col_to_letters(701), "ZZ");
        assert_eq!(col_to_letters(702), "AAA");
    }

    // Test CLI app viewport clamping functions
    #[test]
    fn test_cli_viewport_clamping() {
        // Test vertical clamping
        let mut start_row = -5;
        clamp_viewport_ve(100, &mut start_row);
        assert_eq!(start_row, 0);
        
        start_row = 95;
        clamp_viewport_ve(100, &mut start_row);
        assert_eq!(start_row, 90);
        
        // Test horizontal clamping
        let mut start_col = -5;
        clamp_viewport_hz(100, &mut start_col);
        assert_eq!(start_col, 0);
        
        start_col = 95;
        clamp_viewport_hz(100, &mut start_col);
        assert_eq!(start_col, 90);
    }

    // Process command tests unchanged but now using the actual function
    #[test]
    fn test_cli_process_command() {
        let mut sheet = Box::new(Spreadsheet::new(100, 100));
        let mut status_msg = String::new();
        
        // Test navigation commands
        sheet.top_row = 20;
        sheet.left_col = 20;
        
        process_command(&mut *sheet, "w", &mut status_msg);
        assert_eq!(sheet.top_row, 10);
        
        process_command(&mut *sheet, "s", &mut status_msg);
        assert_eq!(sheet.top_row, 20);
        
        process_command(&mut *sheet, "a", &mut status_msg);
        assert_eq!(sheet.left_col, 10);
        
        process_command(&mut *sheet, "d", &mut status_msg);
        assert_eq!(sheet.left_col, 20);
        
        // Test scroll_to command
        process_command(&mut *sheet, "scroll_to C15", &mut status_msg);
        assert_eq!(sheet.top_row, 14);
        assert_eq!(sheet.left_col, 2);
        
        // Test cell assignment
        process_command(&mut *sheet, "A1=42", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        
        // Test formula assignment
        process_command(&mut *sheet, "B1=A1*2", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 1), 84);
        
        // Test cache clearing
        sheet.cache.insert("test".to_string(), CachedRange {
            value: 42,
            dependencies: HashSet::new(),
        });
        process_command(&mut *sheet, "clear_cache", &mut status_msg);
        assert!(sheet.cache.is_empty());
        
        // Test output control
        assert_eq!(sheet.output_enabled, true);
        process_command(&mut *sheet, "disable_output", &mut status_msg);
        assert_eq!(sheet.output_enabled, false);
        process_command(&mut *sheet, "enable_output", &mut status_msg);
        assert_eq!(sheet.output_enabled, true);
    }

    // Test undo/redo command processing
    #[test]
    #[cfg(feature = "undo_state")]
    fn test_cli_undo_redo_commands() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status_msg = String::new();
        
        // Make changes
        process_command(&mut *sheet, "A1=10", &mut status_msg);
        process_command(&mut *sheet, "A1=20", &mut status_msg);
        
        // Test undo
        process_command(&mut *sheet, "undo", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 10);
        assert!(status_msg.contains("Undo successful"));
        
        // Test redo
        process_command(&mut *sheet, "redo", &mut status_msg);
        assert_eq!(sheet.get_cell_value(0, 0), 20);
        assert!(status_msg.contains("Redo successful"));
    }

    // Test cell history command processing
    #[test]
    #[cfg(feature = "cell_history")]
    fn test_cli_cell_history_command() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status_msg = String::new();
        
        // Make changes to create history
        process_command(&mut *sheet, "A1=10", &mut status_msg);
        process_command(&mut *sheet, "A1=20", &mut status_msg);
        process_command(&mut *sheet, "A1=30", &mut status_msg);
        
        // Test history command - we can only check that it was recognized
        process_command(&mut *sheet, "history A1", &mut status_msg);
        assert_eq!(status_msg, "History displayed");
        assert_eq!(sheet.skip_default_display, true);
    }

    // Test error handling in commands
    #[test]
    fn test_cli_error_handling() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status_msg = String::new();
        
        // Test invalid cell reference
        process_command(&mut *sheet, "Z99=10", &mut status_msg);
        assert_eq!(status_msg, "Cell out of bounds");
        
        // Test invalid cell name
        process_command(&mut *sheet, "ZZZ=10", &mut status_msg);
        assert_eq!(status_msg, "Invalid cell");
        
        // Test invalid command
        process_command(&mut *sheet, "invalid_command", &mut status_msg);
        assert_eq!(status_msg, "unrecognized cmd");
    }
}
