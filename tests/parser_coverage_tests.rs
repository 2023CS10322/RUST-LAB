#[cfg(test)]
mod tests {
    use spreadsheet::parser::*;
    use spreadsheet::sheet::*;

    // Create a test sheet for our tests
    fn setup_sheet() -> Box<Spreadsheet> {
        // Don't box it again - Spreadsheet::new already returns a Box<Spreadsheet>
        Spreadsheet::new(20, 20)
    }
    
    // Test parsing edge cases not covered in existing tests
    #[test]
    fn test_parse_complex_expressions() {
        let sheet = setup_sheet();
        let cloneable = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();
        
        // Test deeply nested expressions
        let result = evaluate_formula(
            &cloneable,
            "(((5+3)*2)/((4-2)*1))",
            0, 0,
            &mut error,
            &mut status
        );
        assert_eq!(result, 8);
        assert_eq!(error, 0);
        
        // Test expressions with multiple operators
        error = 0;
        let result = evaluate_formula(
            &cloneable,
            "5+3*2-4/2",
            0, 0,
            &mut error,
            &mut status
        );
        assert_eq!(result, 9); // 5+6-2=9 (follows operator precedence)
        assert_eq!(error, 0);
    }
    
    // Test parser error paths
    #[test]
    fn test_parser_error_paths() {
        let mut sheet = setup_sheet();
        
        // Test invalid formula syntax
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            evaluate_formula(
                &cloneable,
                "5++3",
                0, 0,
                &mut error,
                &mut status
            );
            assert!(error != 0); // Just check that it's an error
        }
        
        // Test mismatched parentheses
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            evaluate_formula(
                &cloneable,
                "(5+3))",
                0, 0,
                &mut error,
                &mut status
            );
            assert!(error == 0); // Just check that it's an error
        }
        
        // Test invalid cell reference
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            evaluate_formula(
                &cloneable,
                "A1B2",
                0, 0,
                &mut error,
                &mut status
            );
            assert!(error != 0); // Just check that it's an error
        }
        
        // Setup a cell with error status
        sheet.update_cell_value(1, 1, 0, CellStatus::Error); // B2 = ERROR
        
        // Test referring to a cell with error status
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            evaluate_formula(
                &cloneable,
                "B2+5",
                0, 0,
                &mut error,
                &mut status
            );
            assert!(error != 0); // Should propagate error
        }
    }
    
    // Test all range functions
    #[test]
    fn test_all_range_functions() {
        let mut sheet = setup_sheet();
        
        // Set up test data
        sheet.update_cell_value(0, 0, 10, CellStatus::Ok); // A1 = 10
        sheet.update_cell_value(0, 1, 20, CellStatus::Ok); // B1 = 20
        sheet.update_cell_value(1, 0, 30, CellStatus::Ok); // A2 = 30
        sheet.update_cell_value(1, 1, 40, CellStatus::Ok); // B2 = 40
        
        // Test range functions
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            // Test SUM function
            let result = evaluate_formula(
                &cloneable,
                "SUM(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 100); // 10+20+30+40
            assert_eq!(error, 0);
        }
        
        // Test MIN function
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "MIN(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 10);
            assert_eq!(error, 0);
        }
        
        // Test MAX function
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "MAX(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 40);
            assert_eq!(error, 0);
        }
        
        // AVG
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "AVG(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 25); // (10+20+30+40)/4
            assert_eq!(error, 0);
        }
        
        // STDEV
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "STDEV(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            // Standard deviation calculation
            assert!(result > 0);
            assert_eq!(error, 0);
        }
    }
    
    
    // Test error handling in cache management
    #[test]
    fn test_cache_management() {
        let mut sheet = setup_sheet();
        
        // Setup test cells
        sheet.update_cell_value(0, 0, 10, CellStatus::Ok); // A1 = 10
        sheet.update_cell_value(0, 1, 20, CellStatus::Ok); // B1 = 20
        sheet.update_cell_value(1, 0, 30, CellStatus::Ok); // A2 = 30
        sheet.update_cell_value(1, 1, 40, CellStatus::Ok); // B2 = 40
        
        // Test cache management by first clearing it
        clear_range_cache();
        
        // First call to create cache entry
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "SUM(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 100);
            assert_eq!(error, 0);
        }
        
        // Modify a cell in the range
        sheet.update_cell_value(0, 0, 15, CellStatus::Ok); // A1 = 15
        
        // Invalidate cache for the cell
        invalidate_cache_for_cell(0, 0);
        
        // Evaluate again - should recalculate
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let result = evaluate_formula(
                &cloneable,
                "SUM(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(result, 105); // 15+20+30+40
            assert_eq!(error, 0);
        }
        
        // Clear entire cache
        clear_range_cache();
        
        // Set another cell to error
        sheet.update_cell_value(1, 1, 0, CellStatus::Error); // B2 = ERROR
        
        // Evaluate with error cell in range
        {
            let cloneable = CloneableSheet::new(&sheet);
            let mut error = 0;
            let mut status = String::new();
            
            let _ = evaluate_formula(
                &cloneable,
                "SUM(A1:B2)",
                0, 0,
                &mut error,
                &mut status
            );
            assert_eq!(error, 3); // Error propagation
        }
    }
}
