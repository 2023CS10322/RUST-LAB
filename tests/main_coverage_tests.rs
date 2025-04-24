#[cfg(test)]
mod tests {
    use spreadsheet::sheet::*;
    use spreadsheet::parser::*;
    
    // Test functionality without spawning external processes
    #[test]
    fn test_main_dispatcher() {
        // Test using Sheet directly to avoid hanging on GUI/CLI init
        let sheet = Box::new(Spreadsheet::new(10, 10));
        assert_eq!(sheet.total_rows, 10);
        assert_eq!(sheet.total_cols, 10);
        
        // Test core functionality from parser to ensure coverage
        let cloneable = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();
        
        let result = evaluate_formula(&cloneable, "5+5", 0, 0, &mut error, &mut status);
        assert_eq!(result, 10);
        assert_eq!(error, 0);
    }
    
    // Test more complex spreadsheet operations
    #[test]
    fn test_spreadsheet_operations() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status = String::new();
        
        // Set some cells
        sheet.update_cell_formula(0, 0, "42", &mut status);
        sheet.update_cell_formula(0, 1, "A1*2", &mut status);
        
        // Test formula evaluation
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        assert_eq!(sheet.get_cell_value(0, 1), 84);
        
        // Test cache operations
        sheet.cache.clear();
        clear_range_cache();
        
        // Ensure operations still work after cache clear
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        assert_eq!(sheet.get_cell_value(0, 1), 84);
    }
    
    // Test CLI-specific functions if available
    #[test]
    #[cfg(feature = "cli_app")]
    fn test_cli_app_functions() {
        use spreadsheet::cli_app::*;
        
        // Test column letter conversion
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
    }
    
    // Test GUI-specific functions if available
    #[test]
    #[cfg(feature = "gui_app")]
    fn test_gui_app_functions() {
        use spreadsheet::gui_app::*;
        
        // Test column letter conversion
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(25), "Z");
        
        // Test coordinate to cell name conversion
        assert_eq!(coords_to_cell_name(0, 0), "A1");
        assert_eq!(coords_to_cell_name(9, 1), "B10");
    }
}
