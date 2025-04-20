use std::collections::{VecDeque, HashSet, HashMap};

#[derive(PartialEq, Eq, Debug, Clone)]
pub enum CellStatus {
    Ok,
    Error,
}

// Optimize Cell structure by removing redundant fields and using more compact storage
pub struct Cell {
    pub value: i32,
    pub formula_idx: Option<usize>, // Index into formula storage instead of storing entire string
    pub status: CellStatus,
    pub dependencies: HashSet<(i32, i32)>,
    pub dependents: HashSet<(i32, i32)>,
    // Removed row and col fields as they can be derived from the cell's position in the HashMap
}

#[derive(Clone)]
pub struct CachedRange {
    pub value: i32,
    pub dependencies: HashSet<(i32, i32)>,
}

pub struct Spreadsheet {
    pub total_rows: i32,
    pub total_cols: i32,
    pub cells: HashMap<(i32, i32), Cell>, // Sparse representation instead of Vec<Vec<Cell>>
    pub formula_storage: Vec<String>,     // Central storage for all formulas
    pub top_row: i32,
    pub left_col: i32,
    pub output_enabled: bool,
    pub skip_default_display: bool,
    pub cache: HashMap<String, CachedRange>,  // Cached range evaluations
    pub dirty_cells: HashSet<(i32, i32)>,     // Track cells needing recalculation
    pub in_degree: HashMap<(i32, i32), usize>, // For topological sort
}

impl Spreadsheet {
    pub fn get_cell_raw_content(&self, row: i32, col: i32) -> String {
        // Look for the cell in the HashMap using the (row, col) tuple as the key.
        if let Some(cell) = self.cells.get(&(row, col)) {
            // If the cell exists, check if it has a formula index.
            if let Some(idx) = cell.formula_idx {
                // Check if the index is within the bounds of the formula storage.
                // This is a safeguard against potential inconsistencies.
                if idx < self.formula_storage.len() {
                    // If the index is valid, retrieve the string from storage.
                    // Clone it because we need to return an owned String.
                    self.formula_storage[idx].clone()
                } else {
                    // This case indicates an invalid state (index out of bounds).
                    // Log an error and return an empty string.
                    eprintln!(
                        "Error: Cell ({}, {}) has an invalid formula index: {}",
                        row, col, idx
                    );
                    String::new()
                }
            } else {
                // The cell exists in the map, but has no formula_idx.
                // This implies the cell was likely created implicitly (e.g., as a
                // dependency or via get_or_create_cell) but never had a formula
                // explicitly assigned via update_cell_formula.
                // For the purpose of displaying raw content in a formula bar,
                // an empty string is appropriate here.
                String::new()
            }
        } else {
            // If self.cells.get returns None, the cell doesn't exist in our sparse map.
            // Return an empty string to represent an empty cell.
            String::new()
        }}
    pub fn new(rows: i32, cols: i32) -> Box<Spreadsheet> {
        // Create an empty sparse representation instead of initializing all cells
        Box::new(Spreadsheet {
            total_rows: rows,
            total_cols: cols,
            cells: HashMap::new(),
            formula_storage: Vec::new(),
            top_row: 0,
            left_col: 0,
            output_enabled: true,
            skip_default_display: false,
            cache: HashMap::new(),
            dirty_cells: HashSet::new(),
            in_degree: HashMap::new(),
        })
    }

    // Helper method to get or create a cell
    fn get_or_create_cell(&mut self, row: i32, col: i32) -> &mut Cell {
        if !self.cells.contains_key(&(row, col)) {
            self.cells.insert((row, col), Cell {
                value: 0,
                formula_idx: None,
                status: CellStatus::Ok,
                dependencies: HashSet::new(),
                dependents: HashSet::new(),
            });
        }
        self.cells.get_mut(&(row, col)).unwrap()
    }

    
    // Helper method to get cell value (returns 0 for non-existent cells)
    pub fn get_cell_value(&self, row: i32, col: i32) -> i32 {
        self.cells.get(&(row, col)).map_or(0, |cell| cell.value)
    }

    // Helper method to get cell status (returns Ok for non-existent cells)
    pub fn get_cell_status(&self, row: i32, col: i32) -> CellStatus {
        self.cells.get(&(row, col)).map_or(CellStatus::Ok, |cell| cell.status.clone())
    }

    // Helper to get formula string
    pub fn get_formula(&self, row: i32, col: i32) -> Option<String> {
        if let Some(cell) = self.cells.get(&(row, col)) {
            if let Some(idx) = cell.formula_idx {
                return Some(self.formula_storage[idx].clone());
            }
        }
        None
    }

    // Update cell formula (rewritten to use the sparse representation)
    pub fn update_cell_formula(&mut self, row: i32, col: i32, formula: &str, status_msg: &mut String) {
        if valid_formula(self, formula, status_msg) != 0 {
            status_msg.clear();
            status_msg.push_str("Unrecognized");
            return;
        }
        status_msg.clear();
        status_msg.push_str("Ok");

        // First, extract old dependencies
        let old_deps = if let Some(cell) = self.cells.get(&(row, col)) {
            cell.dependencies.clone()
        } else {
            HashSet::new()
        };
        
        // Get old formula
        let old_formula = self.get_formula(row, col);

        // Extract new dependencies
        let new_deps = if !formula.chars().all(|ch| ch.is_digit(10) || ch == '-') {
            extract_dependencies_without_self(formula, self.total_rows, self.total_cols)
        } else {
            HashSet::new()
        };

        // Remove old dependencies
        for &(dep_row, dep_col) in &old_deps {
            if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                if let Some(dep_cell) = self.cells.get_mut(&(dep_row, dep_col)) {
                    dep_cell.dependents.remove(&(row, col));
                }
            }
        }

        // Store the formula centrally and get its index - to avoid borrowing issues
        let formula_idx = {
            // Check if formula already exists to avoid duplication
            let existing_idx = self.formula_storage.iter().position(|f| f == formula);
            match existing_idx {
                Some(idx) => idx,
                None => {
                    let idx = self.formula_storage.len();
                    self.formula_storage.push(formula.to_string());
                    idx
                }
            }
        };

        // Set new formula and clear old dependencies
        {
            let cell = self.get_or_create_cell(row, col);
            cell.dependencies.clear();
            cell.formula_idx = Some(formula_idx);
        }

        // Add new dependencies
        for &(dep_row, dep_col) in &new_deps {
            if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                // Store the dependency in the current cell
                self.get_or_create_cell(row, col).dependencies.insert((dep_row, dep_col));
                
                // Store the current cell as a dependent of the dependency cell
                self.get_or_create_cell(dep_row, dep_col).dependents.insert((row, col));
            }
        }

        // Detect circular dependency
        if has_circular_dependency_by_index(self, row, col) {
            let cell_name = coords_to_cell_name(row, col);
            status_msg.clear();
            status_msg.push_str("Circular dependency detected in cell ");
            status_msg.push_str(&cell_name);

            // Handle old formula index for restoring
            let old_formula_idx = if let Some(old_formula_str) = old_formula {
                // Find or add the old formula back - avoid borrowing issues
                let existing_idx = self.formula_storage.iter().position(|f| f == &old_formula_str);
                match existing_idx {
                    Some(idx) => Some(idx),
                    None => {
                        let idx = self.formula_storage.len();
                        self.formula_storage.push(old_formula_str);
                        Some(idx)
                    }
                }
            } else {
                None
            };
            
            // Now restore the cell's state
            let cell = self.get_or_create_cell(row, col);
            cell.dependencies.clear();
            cell.formula_idx = old_formula_idx;
            
            // Re-add old dependencies
            for &(dep_row, dep_col) in &old_deps {
                if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                    // Add dependency to current cell
                    self.get_or_create_cell(row, col).dependencies.insert((dep_row, dep_col));
                    
                    // Add current cell as dependent to dependency cell
                    self.get_or_create_cell(dep_row, dep_col).dependents.insert((row, col));
                }
            }
            return;
        }

        // Mark this cell as dirty for recalculation
        self.dirty_cells.remove(&(row, col)); 

        // Evaluate the formula
        let mut error_flag = 0;
        let mut s_msg = String::new();
        
        // Create temporary clone for evaluation
        let new_val = {
            let sheet_clone = CloneableSheet::new(self);
            crate::parser::evaluate_formula(&sheet_clone, formula, row, col, &mut error_flag, &mut s_msg)
        };
        
        if error_flag == 3 {
            mark_cell_and_dependents_as_error(self, row, col);
            status_msg.clear();
            status_msg.push_str("Ok");
            return;
        } else if error_flag == 4 {
            status_msg.clear();
            status_msg.push_str("Range out of bounds");
            return;
        } else if error_flag == 1 {
            status_msg.clear();
            status_msg.push_str("Error in formula");
            return;
        } else {
            // Set the value and status first
            {
                let cell = self.get_or_create_cell(row, col);
                cell.value = new_val;
                cell.status = CellStatus::Ok;
            }
            
            // Then get the dependents (to avoid borrowing issues)
            let dependents = if let Some(cell) = self.cells.get(&(row, col)) {
                cell.dependents.clone()
            } else {
                HashSet::new()
            };
            
            // Mark dependent cells as dirty
            for &(dep_row, dep_col) in &dependents {
                self.dirty_cells.insert((dep_row, dep_col));
            }
            
            // Invalidate any cached range functions that depend on this cell
            crate::parser::invalidate_cache_for_cell(row, col);
            
            // Mark dependent cells as dirty more thoroughly
            mark_cell_and_dependents_dirty(self, row, col);
            
            // Use the optimized recalculation
            recalc_affected(self, status_msg);
        }
    }
}

// Utility: converts cell name (e.g. "A1") to (row, col).
pub fn cell_name_to_coords(name: &str) -> Option<(i32, i32)> {
    let mut pos = 0;
    let mut col_val = 0;
    for ch in name.chars() {
        if ch.is_alphabetic() {
            col_val = col_val * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
            pos += 1;
        } else {
            break;
        }
    }
    if col_val == 0 { return None; }
    let col = col_val - 1;
    let mut row_val = 0;
    for ch in name[pos..].chars() {
        if ch.is_digit(10) {
            row_val = row_val * 10 + (ch as i32 - '0' as i32);
        } else {
            return None;
        }
    }
    if row_val <= 0 { return None; }
    Some((row_val - 1, col))
}

// Trims a string in place.
pub fn trim(s: &mut String) {
    *s = s.trim().to_string();
}

// Validates a formula.
pub fn valid_formula(sheet: &Spreadsheet, formula: &str, status_msg: &mut String) -> i32 {
    status_msg.clear();
    let len = formula.len();
    if len == 0 {
        status_msg.push_str("Empty formula");
        return 1;
    }
    if let Some((row, col)) = cell_name_to_coords(formula) {
        if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
            status_msg.push_str("Cell reference out of bounds");
            return 1;
        }
        return 0;
    }
    if formula.trim().parse::<i32>().is_ok() {
        return 0;
    }
        // ── NEW ── Advanced formulas
        if formula.starts_with("IF(") {
            // must have two commas and closing ')'
            let inner = &formula[3..formula.len().saturating_sub(1)];
            if inner.split(',').count() != 3 { status_msg.push_str("IF needs 3 args"); return 1; }
            return 0;
        }
        if formula.starts_with("COUNTIF(") {
            let inner = &formula[8..formula.len().saturating_sub(1)];
            if inner.split(',').count() != 2 { status_msg.push_str("COUNTIF needs 2 args"); return 1; }
            return 0;
        }
        if formula.starts_with("SUMIF(") {
            let inner = &formula[6..formula.len().saturating_sub(1)];
            if inner.split(',').count() != 3 { status_msg.push_str("SUMIF needs 3 args"); return 1; }
            return 0;
        }
        if formula.starts_with("CONCATENATE(") {
            let inner = &formula[12..formula.len().saturating_sub(1)];
            if inner.split(',').count() < 1 { status_msg.push_str("CONCATENATE needs ≥1 arg"); return 1; }
            return 0;
        }
        if formula.starts_with("ROUND(") {
            let inner = &formula[6..formula.len().saturating_sub(1)];
            if inner.split(',').count() != 2 { status_msg.push_str("ROUND needs 2 args"); return 1; }
            return 0;
        }
    
    if formula.starts_with("MAX(") || formula.starts_with("MIN(") ||
       formula.starts_with("SUM(") || formula.starts_with("AVG(") ||
       formula.starts_with("STDEV(")
    {
        let pos = formula.find('(').unwrap_or(0);
        if pos == 0 || formula.chars().nth(pos) != Some('(') {
            status_msg.push_str("Missing '(' after function name");
            return 1;
        }
        if !formula.ends_with(')') {
            status_msg.push_str("Missing closing parenthesis");
            return 1;
        }
        let inner = &formula[pos+1..formula.len()-1];
        let mut inner = inner.trim().to_string();
        if let Some(colon) = inner.find(':') {
            inner.replace_range(colon..colon+1, ":");
            let cell1 = inner[..colon].to_string();
            let cell2 = inner[colon+1..].to_string();
            let mut cell1 = cell1.trim().to_string();
            let mut cell2 = cell2.trim().to_string();
            if cell_name_to_coords(&cell1).is_none() {
                status_msg.push_str("Invalid first cell reference");
                return 1;
            }
            if cell_name_to_coords(&cell2).is_none() {
                status_msg.push_str("Invalid second cell reference");
                return 1;
            }
            let (row1, col1) = cell_name_to_coords(&cell1).unwrap();
            let (row2, col2) = cell_name_to_coords(&cell2).unwrap();
            if row1 < 0 || row1 >= sheet.total_rows || col1 < 0 || col1 >= sheet.total_cols {
                status_msg.push_str("First cell reference out of bounds");
                return 1;
            }
            if row2 < 0 || row2 >= sheet.total_rows || col2 < 0 || col2 >= sheet.total_cols {
                status_msg.push_str("Second cell reference out of bounds");
                return 1;
            }
            if row1 > row2 || col1 > col2 {
                status_msg.push_str("Invalid range order");
                return 1;
            }
            return 0;
        } else {
            status_msg.push_str("Missing colon in range");
            return 1;
        }
    } else if formula.starts_with("SLEEP(") {
        if !formula.ends_with(')') {
            status_msg.push_str("Missing closing parenthesis in SLEEP");
            return 1;
        }
        let inner = &formula[6..formula.len()-1];
        let mut inner = inner.trim().to_string();
        if inner.parse::<i32>().is_ok() {
            return 0;
        } else {
            if cell_name_to_coords(&inner).is_none() {
                status_msg.push_str("Invalid cell reference in SLEEP");
                return 1;
            }
            let (row, col) = cell_name_to_coords(&inner).unwrap();
            if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                status_msg.push_str("Cell reference in up out of bounds");
                return 1;
            }
            return 0;
        }
    }
    let mut op_index = -1;
    let mut i = if formula.starts_with('-') { 1 } else { 0 };
    let chars: Vec<char> = formula.chars().collect();
    while i < chars.len() {
        if chars[i] == '+' || chars[i] == '-' || chars[i] == '*' || chars[i] == '/' {
            op_index = i as i32;
            break;
        }
        i += 1;
    }
    if op_index == -1 {
        status_msg.push_str("Operator not found");
        return 1;
    }
    let left = formula[..op_index as usize].trim();
    let right = formula[op_index as usize + 1..].trim();
    let is_left_int = left.parse::<i32>().is_ok();
    let is_right_int = right.parse::<i32>().is_ok();
    let left_is_cell = cell_name_to_coords(left).is_some();
    let right_is_cell = cell_name_to_coords(right).is_some();
    if (is_left_int || left_is_cell) && (is_right_int || right_is_cell) {
        return 0;
    }
    status_msg.push_str("Invalid formula format");
    1
}

// Optimized: Extract dependencies from a formula using HashSet
pub fn extract_dependencies(sheet: &Spreadsheet, formula: &str) -> HashSet<(i32, i32)> {
    let mut deps: HashSet<(i32, i32)> = HashSet::new();
    let mut p = formula;
    
    while !p.is_empty() {
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() { break; }
            p = &p[ch.len_utf8()..];
        }
        if p.is_empty() { break; }
        
        let start = p;
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() {
                p = &p[ch.len_utf8()..];
            } else { break; }
        }
        while let Some(ch) = p.chars().next() {
            if ch.is_digit(10) {
                p = &p[ch.len_utf8()..];
            } else { break; }
        }
        
        if p.starts_with(':') {
            p = &p[1..];
            let range_start2 = p;
            while let Some(ch) = p.chars().next() {
                if ch.is_alphabetic() {
                    p = &p[ch.len_utf8()..];
                } else { break; }
            }
            while let Some(ch) = p.chars().next() {
                if ch.is_digit(10) {
                    p = &p[ch.len_utf8()..];
                } else { break; }
            }
            
            let len1 = start.find(':').unwrap_or(0);
            let cell_ref1 = &start[..len1];
            let cell_ref2 = &range_start2[..(range_start2.len() - p.len())];
            
            if let (Some((r1, c1)), Some((r2, c2))) =
                (cell_name_to_coords(cell_ref1), cell_name_to_coords(cell_ref2))
            {
                let (start_row, end_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                let (start_col, end_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                
                for rr in start_row..=end_row {
                    for cc in start_col..=end_col {
                        deps.insert((rr, cc));
                    }
                }
            }
        } else {
            let len = start.len() - p.len();
            let cell_ref = &start[..len.min(19)];
            
            if let Some((r, c)) = cell_name_to_coords(cell_ref) {
                deps.insert((r, c));
            }
        }
    }
    
    deps
}

// Detects circular dependency using DFS with HashSets
pub fn has_circular_dependency(sheet: &Spreadsheet, row: i32, col: i32) -> bool {
    let mut visited = HashSet::new();
    let mut stack = vec![(row, col)];
    
    while let Some((r, c)) = stack.pop() {
        visited.insert((r, c));
        
        if let Some(cell) = sheet.cells.get(&(r, c)) {
            for &(dep_row, dep_col) in &cell.dependencies {
                if dep_row == row && dep_col == col {
                    return true;
                }
                
                if !visited.contains(&(dep_row, dep_col)) {
                    stack.push((dep_row, dep_col));
                }
            }
        }
    }
    
    false
}

// Converts (row, col) to cell name.
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

// Optimized: Recalculate affected cells using topological sort with batching
pub fn recalc_affected(sheet: &mut Spreadsheet, status_msg: &mut String) {
    if sheet.dirty_cells.is_empty() {
        return;
    }
    
    // Improved dependency tracking for recalculation
    let dirty_cells = sheet.dirty_cells.clone();
    sheet.dirty_cells.clear(); // Clear before recalculation to allow for new dirty cells
    
    // For large dependency chains, we'll use a more efficient approach
    let mut dependencies_map: HashMap<(i32, i32), HashSet<(i32, i32)>> = HashMap::new();
    let mut in_degree: HashMap<(i32, i32), usize> = HashMap::new();
    let mut to_process = HashSet::new();
    
    // Build the dependency graph more efficiently
    for &(row, col) in &dirty_cells {
        build_dependency_graph(sheet, row, col, &mut dependencies_map, &mut to_process);
    }
    
    // Calculate in-degree for each cell (how many cells it depends on)
    for &node in &to_process {
        in_degree.entry(node).or_insert(0);
    }
    
    for (&cell, deps) in &dependencies_map {
        for &dep in deps {
            if to_process.contains(&dep) {
                *in_degree.entry(dep).or_insert(0) += 1;
            }
        }
    }
    
    // Process in batches for better performance on large chains
    let mut ready_cells: Vec<(i32, i32)> = in_degree.iter()
        .filter(|&(_, &degree)| degree == 0)
        .map(|(&cell, _)| cell)
        .collect();
    
    const BATCH_SIZE: usize = 256; // Process cells in batches for better cache locality
    
    while !ready_cells.is_empty() {
        let batch_end = ready_cells.len().min(BATCH_SIZE);
        let batch = ready_cells.drain(..batch_end).collect::<Vec<_>>();
        
        // Process this batch
        for (row, col) in batch {
            if let Some(formula) = sheet.get_formula(row, col) {
                let mut error_flag = 0;
                let mut s_msg = String::new();
                
                // Create a temporary clone to avoid borrowing issues
                let sheet_clone = CloneableSheet::new(sheet);
                let new_val = crate::parser::evaluate_formula(&sheet_clone, &formula, row, col, &mut error_flag, &mut s_msg);
                
                let cell = sheet.get_or_create_cell(row, col);
                if error_flag == 3 {
                    cell.status = CellStatus::Error;
                    cell.value = 0;
                } else if error_flag != 0 {
                    status_msg.clear();
                    if error_flag == 2 {
                        status_msg.push_str("Invalid range");
                    } else {
                        status_msg.push_str("Error in formula");
                    }
                    return;
                } else {
                    cell.value = new_val;
                    cell.status = CellStatus::Ok;
                }
            }
            
            // Update dependents of this cell and their in-degree
            if let Some(dependents) = dependencies_map.get(&(row, col)) {
                for &dep in dependents {
                    if let Some(deg) = in_degree.get_mut(&dep) {
                        *deg -= 1;
                        if *deg == 0 {
                            ready_cells.push(dep);
                        }
                    }
                }
            }
        }
    }
    
    // Check for cycles (any remaining cells with non-zero in-degree)
    let cells_with_cycles: Vec<(i32, i32)> = in_degree.iter()
        .filter(|&(_, &degree)| degree > 0)
        .map(|(&cell, _)| cell)
        .collect();
    
    // Mark any cells with cycles as errors
    for (row, col) in cells_with_cycles {
        let cell = sheet.get_or_create_cell(row, col);
        cell.status = CellStatus::Error;
        cell.value = 0;
    }
}

// More efficient dependency graph building for large chains
fn build_dependency_graph(
    sheet: &Spreadsheet,
    row: i32, 
    col: i32,
    dependencies_map: &mut HashMap<(i32, i32), HashSet<(i32, i32)>>,
    to_process: &mut HashSet<(i32, i32)>
) {
    let mut queue = VecDeque::new();
    queue.push_back((row, col));
    to_process.insert((row, col));
    
    while let Some((r, c)) = queue.pop_front() {
        // Get dependents (cells that depend on this cell)
        if let Some(cell) = sheet.cells.get(&(r, c)) {
            let dependents = &cell.dependents;
            
            for &(dep_row, dep_col) in dependents {
                // Add this dependency to the map
                dependencies_map.entry((r, c)).or_insert_with(HashSet::new).insert((dep_row, dep_col));
                
                // If this dependent hasn't been processed yet, add it to the queue
                if !to_process.contains(&(dep_row, dep_col)) {
                    to_process.insert((dep_row, dep_col));
                    queue.push_back((dep_row, dep_col));
                }
            }
        }
    }
}

// Extract dependencies without borrowing the sheet - optimized for large formulas
pub fn extract_dependencies_without_self(formula: &str, total_rows: i32, total_cols: i32) -> HashSet<(i32, i32)> {
    // For large range expressions, use a more space-efficient implementation
    // if formula.len() > 100 && formula.contains(':') {
    //     return extract_range_dependencies_optimized(formula, total_rows, total_cols);
    // }
    
    // Original implementation for smaller formulas
    let mut deps: HashSet<(i32, i32)> = HashSet::new();
    let mut p = formula;
    
    while !p.is_empty() {
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() { break; }
            p = &p[ch.len_utf8()..];
        }
        if p.is_empty() { break; }
        
        let start = p;
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() {
                p = &p[ch.len_utf8()..];
            } else { break; }
        }
        while let Some(ch) = p.chars().next() {
            if ch.is_digit(10) {
                p = &p[ch.len_utf8()..];
            } else { break; }
        }
        
        if p.starts_with(':') {
            p = &p[1..];
            let range_start2 = p;
            while let Some(ch) = p.chars().next() {
                if ch.is_alphabetic() {
                    p = &p[ch.len_utf8()..];
                } else { break; }
            }
            while let Some(ch) = p.chars().next() {
                if ch.is_digit(10) {
                    p = &p[ch.len_utf8()..];
                } else { break; }
            }
            
            let len1 = start.find(':').unwrap_or(0);
            let cell_ref1 = &start[..len1];
            let cell_ref2 = &range_start2[..(range_start2.len() - p.len())];
            
            if let (Some((r1, c1)), Some((r2, c2))) =
                (cell_name_to_coords(cell_ref1), cell_name_to_coords(cell_ref2))
            {
                if r1 >= 0 && r1 < total_rows && c1 >= 0 && c1 < total_cols &&
                   r2 >= 0 && r2 < total_rows && c2 >= 0 && c2 < total_cols {
                    let (start_row, end_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                    let (start_col, end_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                    
                    for rr in start_row..=end_row {
                        for cc in start_col..=end_col {
                            deps.insert((rr, cc));
                        }
                    }
                }
            }
        } else {
            let len = start.len() - p.len();
            let cell_ref = &start[..len.min(19)];
            
            if let Some((r, c)) = cell_name_to_coords(cell_ref) {
                if r >= 0 && r < total_rows && c >= 0 && c < total_cols {
                    deps.insert((r, c));
                }
            }
        }
    }
    
    deps
}

// // Optimized extraction for large ranges
// fn extract_range_dependencies_optimized(formula: &str, total_rows: i32, total_cols: i32) -> HashSet<(i32, i32)> {
//     let mut deps = HashSet::new();
    
//     // Fast path for range functions
//     if let Some(open_paren) = formula.find('(') {
//         if let Some(close_paren) = formula.find(')') {
//             let range_part = &formula[open_paren+1..close_paren];
//             if let Some(colon) = range_part.find(':') {
//                 let cell1 = range_part[..colon].trim();
//                 let cell2 = range_part[colon+1..].trim();
                
//                 if let (Some((r1, c1)), Some((r2, c2))) = 
//                     (cell_name_to_coords(cell1), cell_name_to_coords(cell2)) {
                    
//                     if r1 >= 0 && r1 < total_rows && c1 >= 0 && c1 < total_cols &&
//                        r2 >= 0 && r2 < total_rows && c2 >= 0 && c2 < total_cols {
                        
//                         let (start_row, end_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
//                         let (start_col, end_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                        
//                         // For very large ranges, don't materialize all cell references
//                         // if (end_row - start_row + 1) * (end_col - start_col + 1) > 999_999_999 {
//                         //     // Just record the range boundaries to save memory
//                         //     deps.insert((start_row, start_col));
//                         //     deps.insert((start_row, end_col));
//                         //     deps.insert((end_row, start_col));
//                         //     deps.insert((end_row, end_col));
//                         //     // Add some sample cells from the middle
//                         //     deps.insert(((start_row + end_row) / 2, (start_col + end_col) / 2));
//                         //     return deps;
//                         // }
                        
//                         // For smaller ranges, add all cells
//                         for r in start_row..=end_row {
//                             for c in start_col..=end_col {
//                                 deps.insert((r, c));
//                             }
//                         }
//                         return deps;
//                     }
//                 }
//             }
//         }
//     }
    
//     // Fall back to standard extraction for other formulas
//     extract_dependencies_without_self(formula, total_rows, total_cols)
// }

// Marks a cell and its dependents as error
pub fn mark_cell_and_dependents_as_error(sheet: &mut Spreadsheet, row: i32, col: i32) {
    let mut stack = vec![(row, col)];
    let mut visited = HashSet::new();
    
    while let Some((r, c)) = stack.pop() {
        if !visited.insert((r, c)) {
            continue;
        }
        
        let cell = sheet.get_or_create_cell(r, c);
        if cell.status == CellStatus::Error {
            continue;
        }
        
        cell.status = CellStatus::Error;
        cell.value = 0;
        
        let dependents = cell.dependents.clone();
        for &(dep_row, dep_col) in &dependents {
            stack.push((dep_row, dep_col));
        }
    }
}

// Create a cloneable wrapper to avoid borrowing issues
#[derive(Clone)]
pub struct CloneableSheet<'a> {
    sheet: &'a Spreadsheet
}

impl<'a> CloneableSheet<'a> {
    pub fn new(sheet: &'a Spreadsheet) -> Self {
        Self { sheet }
    }
    
    pub fn get_cell(&self, row: i32, col: i32) -> Option<CellView> {
        if row >= 0 && row < self.sheet.total_rows && col >= 0 && col < self.sheet.total_cols {
            if let Some(cell) = self.sheet.cells.get(&(row, col)) {
                return Some(CellView {
                    value: cell.value,
                    status: cell.status.clone(),
                });
            }
            // Return default cell for non-existent cells
            return Some(CellView {
                value: 0,
                status: CellStatus::Ok,
            });
        }
        None
    }
    
    pub fn total_rows(&self) -> i32 {
        self.sheet.total_rows
    }
    
    pub fn total_cols(&self) -> i32 {
        self.sheet.total_cols
    }
}

// Light-weight view of cell data for read-only operations
pub struct CellView {
    pub value: i32,
    pub status: CellStatus,
}

// Detects circular dependency using DFS with indexes to avoid borrowing issues
pub fn has_circular_dependency_by_index(sheet: &Spreadsheet, row: i32, col: i32) -> bool {
    let mut visited = HashSet::new();
    let mut stack = vec![(row, col)];
    
    while let Some((r, c)) = stack.pop() {
        if !visited.insert((r, c)) {
            continue;
        }
        
        // Get dependencies for the current cell
        if let Some(cell) = sheet.cells.get(&(r, c)) {
            // Check for circular dependency
            for &(dep_row, dep_col) in &cell.dependencies {
                if dep_row == row && dep_col == col {
                    return true;
                }
                
                if !visited.contains(&(dep_row, dep_col)) {
                    stack.push((dep_row, dep_col));
                }
            }
        }
    }
    
    false
}

// More memory-efficient dirty cells handling
pub fn mark_cell_and_dependents_dirty(sheet: &mut Spreadsheet, row: i32, col: i32) {
    // For large spreadsheets, avoid excessive memory usage
    const MAX_DIRTY_CELLS: usize = 1000000;
    
    // if sheet.dirty_cells.len() > MAX_DIRTY_CELLS {
    //     // If we have too many dirty cells already, do a partial recalculation now
    //     let status_msg = &mut String::new();
    //     recalc_affected(sheet, status_msg);
    // }
    
    let mut queue = VecDeque::new();
    let mut visited = HashSet::new();
    
    // Start with the immediate dependents
    let dependents = if let Some(cell) = sheet.cells.get(&(row, col)) {
        cell.dependents.clone()
    } else {
        HashSet::new()
    };
    
    for &(dep_row, dep_col) in &dependents {
        queue.push_back((dep_row, dep_col));
    }
    
    // Process the dependency graph with a limit to avoid stack overflows
    let mut cells_visited = 0;
    const MAX_VISIT: usize = 5000;
    
    while let Some((r, c)) = queue.pop_front() {
        cells_visited += 1;
        // if cells_visited > MAX_VISIT {
        //     // Too many cells in the chain, mark them for batch processing later
        //     sheet.dirty_cells.insert((r, c));
        //     continue;
        // }
        
        if !visited.insert((r, c)) {
            continue; // Already visited
        }
        
        // Mark as dirty
        sheet.dirty_cells.insert((r, c));
        
        // Invalidate any range functions that depend on this cell
        crate::parser::invalidate_cache_for_cell(r, c);
        
        // Add this cell's dependents to the queue (with batch processing for large chains)
        let next_dependents = if let Some(cell) = sheet.cells.get(&(r, c)) {
            cell.dependents.clone()
        } else {
            HashSet::new()
        };

        for &(dep_row, dep_col) in &next_dependents {
            sheet.dirty_cells.insert((dep_row, dep_col));
        }

        // if next_dependents.len() > 1000000 {
        //     // For cells with many dependents, just mark them all as dirty without full traversal
        //     for &(dep_row, dep_col) in &next_dependents {
        //         sheet.dirty_cells.insert((dep_row, dep_col));
        //     }
        // } else {
        //     // For cells with fewer dependents, continue normal traversal
        //     for &(dep_row, dep_col) in &next_dependents {
        //         if !visited.contains(&(dep_row, dep_col)) {
        //             queue.push_back((dep_row, dep_col));
        //         }
        //     }
        // }
    }
}
