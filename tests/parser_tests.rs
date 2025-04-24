#[cfg(test)]
mod tests {
    use spreadsheet::parser::*;
    use spreadsheet::sheet::*;

    // Helper function to create a test spreadsheet
    fn setup_test_sheet() -> Box<Spreadsheet> {
        let mut sheet = Spreadsheet::new(10, 10);
        sheet.update_cell_value(0, 0, 42, CellStatus::Ok); // A1 = 42
        sheet.update_cell_value(0, 1, 24, CellStatus::Ok); // B1 = 24
        sheet.update_cell_value(1, 0, 10, CellStatus::Ok); // A2 = 10
        sheet.update_cell_value(1, 1, 20, CellStatus::Ok); // B2 = 20
        sheet
    }

    #[test]
    fn test_basic_arithmetic() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test basic arithmetic operations
        assert_eq!(evaluate_formula(&cloneable_sheet, "42", 0, 0, &mut error, &mut status), 42);
        assert_eq!(evaluate_formula(&cloneable_sheet, "2+3", 0, 0, &mut error, &mut status), 5);
        assert_eq!(evaluate_formula(&cloneable_sheet, "10-3", 0, 0, &mut error, &mut status), 7);
        assert_eq!(evaluate_formula(&cloneable_sheet, "4*5", 0, 0, &mut error, &mut status), 20);
        assert_eq!(evaluate_formula(&cloneable_sheet, "15/3", 0, 0, &mut error, &mut status), 5);
        assert_eq!(error, 0);
    }

    #[test]
    fn test_cell_references() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test cell references
        assert_eq!(evaluate_formula(&cloneable_sheet, "A1", 0, 0, &mut error, &mut status), 42);
        assert_eq!(evaluate_formula(&cloneable_sheet, "B1", 0, 0, &mut error, &mut status), 24);
        assert_eq!(evaluate_formula(&cloneable_sheet, "A1+B1", 0, 0, &mut error, &mut status), 66);
        assert_eq!(error, 0);
    }

    #[test]
    fn test_range_functions() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test range functions
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUM(A1:B2)", 0, 0, &mut error, &mut status), 96);
        assert_eq!(evaluate_formula(&cloneable_sheet, "MIN(A1:B2)", 0, 0, &mut error, &mut status), 10);
        assert_eq!(evaluate_formula(&cloneable_sheet, "MAX(A1:B2)", 0, 0, &mut error, &mut status), 42);
        assert_eq!(evaluate_formula(&cloneable_sheet, "AVG(A1:B2)", 0, 0, &mut error, &mut status), 24);
        assert_eq!(error, 0);
    }

    #[test]
    fn test_empty_and_whitespace() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        evaluate_formula(&cloneable_sheet, "", 0, 0, &mut error, &mut status);
        assert_eq!(error, 1);
        
        error = 0;
        status.clear();
        evaluate_formula(&cloneable_sheet, "   ", 0, 0, &mut error, &mut status);
        assert_eq!(error, 1);
    }

    #[test]
    fn test_complex_arithmetic() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cloneable_sheet, "(2+3)*4", 0, 0, &mut error, &mut status), 20);
        assert_eq!(evaluate_formula(&cloneable_sheet, "10/(2+3)", 0, 0, &mut error, &mut status), 2);
        assert_eq!(evaluate_formula(&cloneable_sheet, "-5+10", 0, 0, &mut error, &mut status), 5);
        assert_eq!(error, 0);
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formulas() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test IF function
        assert_eq!(evaluate_formula(&cloneable_sheet, "IF(A1>B1,1,0)", 0, 0, &mut error, &mut status), 1);
        assert_eq!(evaluate_formula(&cloneable_sheet, "IF(A1<B1,1,0)", 0, 0, &mut error, &mut status), 0);

        // Test COUNTIF function
        assert_eq!(evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\">20\")", 0, 0, &mut error, &mut status), 2);

        // Test SUMIF function
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUMIF(A1:B2,\">20\",A1:B2)", 0, 0, &mut error, &mut status), 66);
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formulas_comparison() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test comparison operators
        assert_eq!(evaluate_formula(&cloneable_sheet, "A1>B1", 0, 0, &mut error, &mut status), 1); // 42 > 24
        assert_eq!(evaluate_formula(&cloneable_sheet, "A1>=24", 0, 0, &mut error, &mut status), 1); // 42 >= 24
        assert_eq!(evaluate_formula(&cloneable_sheet, "B1<A1", 0, 0, &mut error, &mut status), 1); // 24 < 42
        assert_eq!(evaluate_formula(&cloneable_sheet, "B1<=A1", 0, 0, &mut error, &mut status), 1); // 24 <= 42
        assert_eq!(evaluate_formula(&cloneable_sheet, "B1==24", 0, 0, &mut error, &mut status), 1); // 24 == 24
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formulas_countif() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test COUNTIF with different conditions
        assert_eq!(evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\">20\")", 0, 0, &mut error, &mut status), 2);
        assert_eq!(evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\"<20\")", 0, 0, &mut error, &mut status), 1);
        assert_eq!(evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\"=24\")", 0, 0, &mut error, &mut status), 1);
        assert_eq!(evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\">=20\")", 0, 0, &mut error, &mut status), 3);
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formulas_sumif() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test SUMIF with different conditions
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUMIF(A1:B2,\">20\",A1:B2)", 0, 0, &mut error, &mut status), 66);
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUMIF(A1:B2,\"<20\",A1:B2)", 0, 0, &mut error, &mut status), 10);
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUMIF(A1:B2,\"=24\",A1:B2)", 0, 0, &mut error, &mut status), 24);
    }

    #[test]
    fn test_complex_range_functions() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test nested range functions
        assert_eq!(evaluate_formula(&cloneable_sheet, "SUM(A1:B1)/2", 0, 0, &mut error, &mut status), 33); // (42+24)/2
        assert_eq!(evaluate_formula(&cloneable_sheet, "MAX(A1:B1)-MIN(A1:B1)", 0, 0, &mut error, &mut status), 18); // 42-24
        assert_eq!(evaluate_formula(&cloneable_sheet, "AVG(A1:B2)*2", 0, 0, &mut error, &mut status), 48); // 24*2
    }

    #[test]
    fn test_cache_behavior() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // First evaluation should cache
        let result1 = evaluate_formula(&cloneable_sheet, "SUM(A1:B2)", 0, 0, &mut error, &mut status);
        
        // Clear cache
        clear_range_cache();
        
        // Second evaluation should give same result
        let result2 = evaluate_formula(&cloneable_sheet, "SUM(A1:B2)", 0, 0, &mut error, &mut status);
        
        assert_eq!(result1, result2);
        assert_eq!(result1, 96);
    }

    #[test]
    fn test_formula_parsing() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test numeric literals
        assert_eq!(evaluate_formula(&cloneable_sheet, "42", 0, 0, &mut error, &mut status), 42);
        assert_eq!(error, 0);

        // Test negative numbers
        error = 0;
        assert_eq!(evaluate_formula(&cloneable_sheet, "-42", 0, 0, &mut error, &mut status), -42);
        assert_eq!(error, 0);

        // Test parenthesized expressions
        error = 0;
        assert_eq!(evaluate_formula(&cloneable_sheet, "(2+3)", 0, 0, &mut error, &mut status), 5);
        assert_eq!(error, 0);

        // Test complex expressions
        error = 0;
        assert_eq!(evaluate_formula(&cloneable_sheet, "(10+5)*2", 0, 0, &mut error, &mut status), 30);
        assert_eq!(evaluate_formula(&cloneable_sheet, "20/(5-3)", 0, 0, &mut error, &mut status), 10);
        assert_eq!(error, 0);
    }

    #[test]
    fn test_sleep_functionality() {
        let mut sheet = setup_test_sheet();  // Make sheet mutable
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test sleep with literal
        let start = std::time::Instant::now();
        let result = evaluate_formula(&cloneable_sheet, "SLEEP(1)", 0, 0, &mut error, &mut status);
        let elapsed = start.elapsed();
        assert!(elapsed.as_secs() >= 1);
        assert_eq!(result, 1);
        assert_eq!(error, 0);

        // Test sleep with cell reference
        error = 0;
        sheet.update_cell_value(0, 0, 0, CellStatus::Ok);  // Now works because sheet is mutable
        let cloneable_sheet = CloneableSheet::new(&sheet);  // Create new cloneable after mutation
        let result = evaluate_formula(&cloneable_sheet, "SLEEP(A1)", 0, 0, &mut error, &mut status);
        assert_eq!(result, 0);
        assert_eq!(error, 0);

        // Test negative sleep value
        error = 0;
        let result = evaluate_formula(&cloneable_sheet, "SLEEP(-1)", 0, 0, &mut error, &mut status);
        assert_eq!(result, -1);
        assert_eq!(error, 0);
    }

    #[test]
    fn test_division_by_zero() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        evaluate_formula(&cloneable_sheet, "10/0", 0, 0, &mut error, &mut status);
        assert_eq!(error, 3); // Division by zero should set error to 3

        error = 0;
        evaluate_formula(&cloneable_sheet, "5/(2-2)", 0, 0, &mut error, &mut status);
        assert_eq!(error, 3); // Division by zero via expression
    }

    #[test]
    fn test_out_of_bounds_references() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test referencing cell outside sheet boundaries
        evaluate_formula(&cloneable_sheet, "Z99", 0, 0, &mut error, &mut status);
        assert_eq!(error, 4); // Should set error code 4 for out of bounds
    }

    #[test]
    fn test_error_propagation() {
        let mut sheet = setup_test_sheet();
        // Create error cell
        sheet.update_cell_value(2, 0, 0, CellStatus::Error); // A3 = ERROR
        
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Reference to error cell should propagate error
        evaluate_formula(&cloneable_sheet, "A3+1", 0, 0, &mut error, &mut status);
        assert_eq!(error, 3); // Error propagation
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formulas_nested() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test nested IF functions
        assert_eq!(
            evaluate_formula(
                &cloneable_sheet,
                "IF(A1>30,IF(B1>20,1,2),3)",
                0,
                0,
                &mut error,
                &mut status
            ),
            1
        ); // A1=42>30 and B1=24>20, so first condition true, second true, result=1
    }

    #[test]
    fn test_range_functions_with_errors() {
        let mut sheet = setup_test_sheet();
        // Create error cell within range
        sheet.update_cell_value(0, 2, 0, CellStatus::Error); // C1 = ERROR
        
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Range containing error cell should propagate error
        evaluate_formula(&cloneable_sheet, "SUM(A1:C1)", 0, 0, &mut error, &mut status);
        assert_eq!(error, 3);
    }

    #[test]
    fn test_range_functions_empty_range() {
        let sheet = Spreadsheet::new(10, 10);
        // All cells are empty/zero
        
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test AVG on empty range
        // Even with empty cells, the range itself is valid - just contains zeros
        let result = evaluate_formula(&cloneable_sheet, "AVG(A1:B2)", 0, 0, &mut error, &mut status);
        assert_eq!(error, 0);  // No error for empty range with zeros
        assert_eq!(result, 0);  // Result should be 0
        
        // Test STDEV on empty range
        error = 0;
        let result = evaluate_formula(&cloneable_sheet, "STDEV(A1:B2)", 0, 0, &mut error, &mut status);
        assert_eq!(error, 0);
        assert_eq!(result, 0);  // Standard deviation of all zeros is 0
    }

    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_advanced_formula_errors() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test IF with missing arguments
        evaluate_formula(&cloneable_sheet, "IF(A1>B1,1)", 0, 0, &mut error, &mut status);
        assert_eq!(error, 1);

        // Test COUNTIF with invalid condition
        error = 0;
        evaluate_formula(&cloneable_sheet, "COUNTIF(A1:B2,\">abc\")", 0, 0, &mut error, &mut status);
        assert_eq!(error, 1);
    }

    #[test]
    fn test_cache_invalidation() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // First call caches
        let _result1 = evaluate_formula(&cloneable_sheet, "SUM(A1:B2)", 0, 0, &mut error, &mut status);
        
        // Invalidate specific cell
        invalidate_cache_for_cell(0, 0); // Invalidate A1
        
        // Verify cache still works for other cells
        let result2 = evaluate_formula(&cloneable_sheet, "SUM(B1:B2)", 0, 0, &mut error, &mut status);
        assert_eq!(result2, 44); // B1(24) + B2(20)
        
        // Clear entire cache
        clear_range_cache();
        
        // Instead of directly accessing the cache, verify it was cleared by checking behavior
        // When cache is cleared, evaluating the same formula again should recompute
        let _called_count = 0;
        
        // Create a sheet where we can modify a value to ensure recalculation
        let mut new_sheet = setup_test_sheet();
        new_sheet.update_cell_value(0, 0, 100, CellStatus::Ok); // Change A1 to 100
        
        let new_cloneable = CloneableSheet::new(&new_sheet);
        let result_after_clear = evaluate_formula(&new_cloneable, "SUM(A1:B2)", 0, 0, &mut error, &mut status);
        
        // Should reflect the updated A1 value (100+24+10+20=154)
        assert_eq!(result_after_clear, 154);
    }

    #[test]
    fn test_complex_nested_expressions() {
        let sheet = setup_test_sheet();
        let cloneable_sheet = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Test complex nested expression
        assert_eq!(
            evaluate_formula(
                &cloneable_sheet,
                "(A1 + B1) * (A2 - B2/2)",
                0,
                0,
                &mut error,
                &mut status
            ),
            66 * 0 // (42+24) * (10-20/2) = 66 * 0 = 0
        );
        assert_eq!(error, 0);
    }
}