use std::collections::{VecDeque, HashSet, HashMap};

#[derive(PartialEq, Eq, Debug, Clone)]
pub enum CellStatus {
    Ok,
    Error,
}

pub struct Cell {
    pub value: i32,
    pub formula: Option<String>,
    pub status: CellStatus,
    pub dependencies: HashSet<(u16, u16)>,  // (row, col) indices
    pub dependents: HashSet<(u16, u16)>,
    pub row: u16,
    pub col: u16,
}

#[derive(Clone)]
pub struct CachedRange {
    pub value: i32,
    pub dependencies: HashSet<(u16, u16)>,
}

pub struct Spreadsheet {
    pub total_rows: u16,
    pub total_cols: u16,
    pub cells: Vec<Vec<Cell>>,
    pub top_row: u16,
    pub left_col: u16,
    pub output_enabled: bool,
    pub skip_default_display: bool,
    pub cache: HashMap<String, CachedRange>,  // Cached range evaluations
    pub dirty_cells: HashSet<(u16, u16)>,     // Track cells needing recalculation
    pub in_degree: HashMap<(u16, u16), usize>, // For topological sort
}

impl Spreadsheet {
    pub fn new(rows: u16, cols: u16) -> Box<Spreadsheet> {
        let mut cells = Vec::with_capacity(rows as usize);
        for r in 0..rows {
            let mut row_cells = Vec::with_capacity(cols as usize);
            for c in 0..cols {
                row_cells.push(Cell {
                    value: 0,
                    formula: None,
                    status: CellStatus::Ok,
                    dependencies: HashSet::new(),
                    dependents: HashSet::new(),
                    row: r,
                    col: c,
                });
            }
            cells.push(row_cells);
        }

        Box::new(Spreadsheet {
            total_rows: rows,
            total_cols: cols,
            cells,
            top_row: 0,
            left_col: 0,
            output_enabled: true,
            skip_default_display: false,
            cache: HashMap::new(),
            dirty_cells: HashSet::new(),
            in_degree: HashMap::new(),
        })
    }

    // Update cell formula (rewritten to use HashSets)
    pub fn update_cell_formula(&mut self, row: u16, col: u16, formula: &str, status_msg: &mut String) {
        if valid_formula(self, formula, status_msg) != 0 {
            status_msg.clear();
            status_msg.push_str("Unrecognized");
            return;
        }
        status_msg.clear();
        status_msg.push_str("Ok");

        // First, extract all necessary info to avoid multiple borrows
        let old_deps = {
            let cell = &self.cells[row as usize][col as usize];
            cell.dependencies.clone()
        };
        
        let old_formula = {
            let cell = &self.cells[row as usize][col as usize];
            cell.formula.clone()
        };

        // Extract new dependencies
        let new_deps = if !formula.chars().all(|ch| ch.is_digit(10) || ch == '-') {
            extract_dependencies_without_self(formula, self.total_rows, self.total_cols)
        } else {
            HashSet::new()
        };

        // Remove old dependencies
        for &(dep_row, dep_col) in &old_deps {
            if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                let dep_cell = &mut self.cells[dep_row as usize][dep_col as usize];
                dep_cell.dependents.remove(&(row, col));
            }
        }

        // Set new formula and clear old dependencies
        {
            let cell = &mut self.cells[row as usize][col as usize];
            cell.dependencies.clear();
            cell.formula = Some(formula.to_string());
        }

        // Add new dependencies - modified to avoid multiple mutable borrows
        for &(dep_row, dep_col) in &new_deps {
            if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                // First, store the dependency in the current cell
                self.cells[row as usize][col as usize].dependencies.insert((dep_row, dep_col));
                
                // Then, store the current cell as a dependent of the dependency cell
                self.cells[dep_row as usize][dep_col as usize].dependents.insert((row, col));
            }
        }

        // Detect circular dependency
        if has_circular_dependency_by_index(self, row, col) {
            let cell_name = coords_to_cell_name(row, col);
            status_msg.clear();
            status_msg.push_str("Circular dependency detected in cell ");
            status_msg.push_str(&cell_name);

            // Clear dependencies and restore formula - avoid multiple mutable borrows
            self.cells[row as usize][col as usize].dependencies.clear();
            self.cells[row as usize][col as usize].formula = old_formula;
            
            // Re-add old dependencies - fixed to avoid mutable borrow issues
            for &(dep_row, dep_col) in &old_deps {
                if dep_row >= 0 && dep_row < self.total_rows && dep_col >= 0 && dep_col < self.total_cols {
                    // Add dependency to current cell
                    self.cells[row as usize][col as usize].dependencies.insert((dep_row, dep_col));
                    
                    // Add current cell as dependent to dependency cell
                    self.cells[dep_row as usize][dep_col as usize].dependents.insert((row, col));
                }
            }
            return;
        }

        // Mark this cell as dirty for recalculation
        self.dirty_cells.insert((row, col));

        // Evaluate the formula
        let mut error_flag = 0;
        let mut s_msg = String::new();
        
        // Create temporary clone for immutable reference to sheet
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
            let cell = &mut self.cells[row as usize][col as usize];
            cell.value = new_val;
            cell.status = CellStatus::Ok;
            
            // Mark dependent cells as dirty
            for &(dep_row, dep_col) in &cell.dependents {
                self.dirty_cells.insert((dep_row, dep_col));
            }
            
            // After updating the cell value, invalidate any cached range functions 
            // that depend on this cell
            crate::parser::invalidate_cache_for_cell(row, col);
            
            // Mark dependent cells as dirty more thoroughly
            mark_cell_and_dependents_dirty(self, row, col);
            
            // Use the optimized recalculation
            recalc_affected(self, status_msg);
        }
    }
}

// Utility: converts cell name (e.g. "A1") to (row, col).
pub fn cell_name_to_coords(name: &str) -> Option<(u16, u16)> {
    let mut pos = 0;
    let mut col_val = 0;
    for ch in name.chars() {
        if ch.is_alphabetic() {
            col_val = col_val * 26 + (ch.to_ascii_uppercase() as u16 - 'A' as u16 + 1);
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
            row_val = row_val * 10 + (ch as u16 - '0' as u16);
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
pub fn valid_formula(sheet: &Spreadsheet, formula: &str, status_msg: &mut String) -> u16 {
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
    if formula.trim().parse::<u16>().is_ok() {
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
        if inner.parse::<u16>().is_ok() {
            return 0;
        } else {
            if cell_name_to_coords(&inner).is_none() {
                status_msg.push_str("Invalid cell reference in SLEEP");
                return 1;
            }
            let (row, col) = cell_name_to_coords(&inner).unwrap();
            if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                status_msg.push_str("Cell reference in SLEEP out of bounds");
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
pub fn extract_dependencies(sheet: &Spreadsheet, formula: &str) -> HashSet<(u16, u16)> {
    let mut deps: HashSet<(u16, u16)> = HashSet::new();
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
pub fn has_circular_dependency(sheet: &Spreadsheet, row: u16, col: u16) -> bool {
    let mut visited = HashSet::new();
    let mut stack = vec![(row, col)];
    
    while let Some((r, c)) = stack.pop() {
        visited.insert((r, c));
        
        let cell = &sheet.cells[r as usize][c as usize];
        for &(dep_row, dep_col) in &cell.dependencies {
            if dep_row == row && dep_col == col {
                return true;
            }
            
            if !visited.contains(&(dep_row, dep_col)) {
                stack.push((dep_row, dep_col));
            }
        }
    }
    
    false
}

// Converts (row, col) to cell name.
pub fn coords_to_cell_name(row: u16, col: u16) -> String {
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

// Optimized: Recalculate affected cells using topological sort
pub fn recalc_affected(sheet: &mut Spreadsheet, status_msg: &mut String) {
    if sheet.dirty_cells.is_empty() {
        return;
    }
    
    // Improved dependency tracking for recalculation
    let dirty_cells = sheet.dirty_cells.clone();
    sheet.dirty_cells.clear(); // Clear before recalculation to allow for new dirty cells
    
    // Build the full dependency graph for affected cells
    let mut graph = HashMap::new();
    let mut to_process = HashSet::new();
    
    // Start with dirty cells
    for &(row, col) in &dirty_cells {
        to_process.insert((row, col));
        graph.entry((row, col)).or_insert_with(Vec::new);
        
        // Also add all cells in the dependency chain
        let mut queue = VecDeque::new();
        queue.push_back((row, col));
        let mut visited = HashSet::new();
        visited.insert((row, col));
        
        while let Some((r, c)) = queue.pop_front() {
            // Add dependents
            let dependents = sheet.cells[r as usize][c as usize].dependents.clone();
            for &(dep_row, dep_col) in &dependents {
                if !visited.contains(&(dep_row, dep_col)) {
                    visited.insert((dep_row, dep_col));
                    to_process.insert((dep_row, dep_col));
                    queue.push_back((dep_row, dep_col));
                    
                    // Add edge in the graph
                    graph.entry((r, c)).or_insert_with(Vec::new).push((dep_row, dep_col));
                    graph.entry((dep_row, dep_col)).or_insert_with(Vec::new);
                }
            }
        }
    }
    
    // Calculate in-degree for each cell in the graph
    let mut in_degree = HashMap::new();
    for &node in &to_process {
        in_degree.insert(node, 0);
    }
    
    for (_, edges) in &graph {
        for &edge in edges {
            *in_degree.entry(edge).or_insert(0) += 1;
        }
    }
    
    // Queue for topological sort
    let mut queue = VecDeque::new();
    for &node in &to_process {
        if in_degree[&node] == 0 {
            queue.push_back(node);
        }
    }
    
    // Process cells in topological order
    while let Some((row, col)) = queue.pop_front() {
        let formula_option = sheet.cells[row as usize][col as usize].formula.clone();
        
        if let Some(formula) = formula_option {
            let mut error_flag = 0;
            let mut s_msg = String::new();
            
            // Create a temporary clone to avoid borrowing issues
            let sheet_clone = CloneableSheet::new(sheet);
            let new_val = crate::parser::evaluate_formula(&sheet_clone, &formula, row, col, &mut error_flag, &mut s_msg);
            
            let cell = &mut sheet.cells[row as usize][col as usize];
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
        
        // Update in-degree for dependent cells
        if let Some(edges) = graph.get(&(row, col)) {
            for &edge in edges {
                *in_degree.get_mut(&edge).unwrap() -= 1;
                if in_degree[&edge] == 0 {
                    queue.push_back(edge);
                }
            }
        }
    }
}

// Marks a cell and its dependents as error
pub fn mark_cell_and_dependents_as_error(sheet: &mut Spreadsheet, row: u16, col: u16) {
    let mut stack = vec![(row, col)];
    let mut visited = HashSet::new();
    
    while let Some((r, c)) = stack.pop() {
        if !visited.insert((r, c)) {
            continue;
        }
        
        let cell = &mut sheet.cells[r as usize][c as usize];
        if cell.status == CellStatus::Error {
            continue;
        }
        
        cell.status = CellStatus::Error;
        cell.value = 0;
        
        for &(dep_row, dep_col) in &cell.dependents {
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
    
    pub fn get_cell(&self, row: u16, col: u16) -> Option<&Cell> {
        if row >= 0 && row < self.sheet.total_rows && col >= 0 && col < self.sheet.total_cols {
            Some(&self.sheet.cells[row as usize][col as usize])
        } else {
            None
        }
    }
    
    pub fn total_rows(&self) -> u16 {
        self.sheet.total_rows
    }
    
    pub fn total_cols(&self) -> u16 {
        self.sheet.total_cols
    }
}

// Extract dependencies without borrowing the sheet
pub fn extract_dependencies_without_self(formula: &str, total_rows: u16, total_cols: u16) -> HashSet<(u16, u16)> {
    let mut deps: HashSet<(u16, u16)> = HashSet::new();
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

// Detects circular dependency using DFS with indexes to avoid borrowing issues
pub fn has_circular_dependency_by_index(sheet: &Spreadsheet, row: u16, col: u16) -> bool {
    let mut visited = HashSet::new();
    let mut stack = vec![(row, col)];
    
    while let Some((r, c)) = stack.pop() {
        if !visited.insert((r, c)) {
            continue;
        }
        
        // Create local copy of dependencies to avoid borrowing issues
        let deps = sheet.cells[r as usize][c as usize].dependencies.clone();
        
        for &(dep_row, dep_col) in &deps {
            if dep_row == row && dep_col == col {
                return true;
            }
            
            if !visited.contains(&(dep_row, dep_col)) {
                stack.push((dep_row, dep_col));
            }
        }
    }
    
    false
}

// Add a function to recursively mark all dependent cells as dirty
pub fn mark_cell_and_dependents_dirty(sheet: &mut Spreadsheet, row: u16, col: u16) {
    let mut queue = VecDeque::new();
    let mut visited = HashSet::new();
    
    // Start with the immediate dependents
    {
        let cell = &sheet.cells[row as usize][col as usize];
        for &(dep_row, dep_col) in &cell.dependents {
            queue.push_back((dep_row, dep_col));
        }
    }
    
    // Process the entire dependency graph
    while let Some((r, c)) = queue.pop_front() {
        if !visited.insert((r, c)) {
            continue; // Already visited
        }
        
        // Mark as dirty
        sheet.dirty_cells.insert((r, c));
        
        // Invalidate any range functions that depend on this cell
        crate::parser::invalidate_cache_for_cell(r, c);
        
        // Add this cell's dependents to the queue
        let dependents = sheet.cells[r as usize][c as usize].dependents.clone();
        for &(dep_row, dep_col) in &dependents {
            if !visited.contains(&(dep_row, dep_col)) {
                queue.push_back((dep_row, dep_col));
            }
        }
    }
}
