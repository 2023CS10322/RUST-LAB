#[cfg(test)]
mod tests {
    use spreadsheet::sheet::*;


    #[test]
    fn test_edge_cases_in_dependency_tracking() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status = String::new();
        
        // Test circular dependency detection
        sheet.update_cell_formula(0, 0, "10", &mut status); // A1 = 10
        sheet.update_cell_formula(0, 1, "A1+5", &mut status); // B1 = A1+5 = 15
        sheet.update_cell_formula(0, 0, "B1+5", &mut status); // A1 = B1+5 (circular)
        
        // Should detect circular dependency and reject the formula
        assert!(status.contains("Circular dependency"));
        
        // A1 should keep its original value
        assert_eq!(sheet.get_cell_value(0, 0), 10);
        
        // Test dependency chain with error propagation
        sheet.update_cell_formula(1, 0, "10", &mut status); // A2 = 10
        sheet.update_cell_formula(1, 1, "A2/0", &mut status); // B2 = A2/0 (error)
        sheet.update_cell_formula(1, 2, "B2+5", &mut status); // C2 references cell with error
        
        // B2 should have error status
        assert_eq!(sheet.get_cell_status(1, 1), CellStatus::Error);
        
        // C2 should also have error status (propagated)
        assert_eq!(sheet.get_cell_status(1, 2), CellStatus::Error);
    }
    
    #[test]
    fn test_complex_formula_storage() {
        let mut sheet = Box::new(Spreadsheet::new(10, 10));
        let mut status = String::new();
        
        // Add a formula
        sheet.update_cell_formula(0, 0, "10+20", &mut status);
        
        // Get the formula back
        let formula = sheet.get_formula(0, 0);
        assert_eq!(formula, Some("10+20".to_string()));
        
        // Add the same formula again to test formula deduplication
        sheet.update_cell_formula(1, 0, "10+20", &mut status);
        
        // Get formula for second cell
        let formula2 = sheet.get_formula(1, 0);
        assert_eq!(formula2, Some("10+20".to_string()));
        
        // Check that they have the same formula index
        let cell1 = sheet.cells.get(&(0, 0)).unwrap();
        let cell2 = sheet.cells.get(&(1, 0)).unwrap();
        assert_eq!(cell1.formula_idx, cell2.formula_idx);
        
        // Raw content should return the formula
        let raw1 = sheet.get_cell_raw_content(0, 0);
        assert_eq!(raw1, "10+20");
    }
}
