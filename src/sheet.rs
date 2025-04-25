#![allow(warnings)]
use std::collections::{HashMap, HashSet, VecDeque};

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
    // --- Additions for Cell History ---
    #[cfg(feature = "cell_history")]
    pub history: VecDeque<i32>, // Store last N values
                                // --- End Additions ---
                                // Removed row and col fields as they can be derived from the cell's position in the HashMap
}

// --- Additions for Undo State ---
#[cfg(feature = "undo_state")]
#[derive(Clone, Debug)] // Clone might be useful, Debug for inspection
struct PreviousCellState {
    row: i32,
    col: i32,
    previous_formula_idx: Option<usize>, // Store index directly
    previous_value: i32,
    previous_status: CellStatus,
    previous_dependencies: HashSet<(i32, i32)>,
    // Store the dependents that pointed *to this cell* before the change
    previous_dependents_of_cell: HashSet<(i32, i32)>,
}
// --- End Additions ---

// Helper constant for history size
#[cfg(feature = "cell_history")]
const MAX_HISTORY_SIZE: usize = 10;

// --- Define the maximum number of undo levels ---
#[cfg(feature = "undo_state")]
const MAX_UNDO_LEVELS: usize = 10; // Set the desired history limit [User Requirement]

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
    pub cache: HashMap<String, CachedRange>, // Cached range evaluations
    pub dirty_cells: HashSet<(i32, i32)>,    // Track cells needing recalculation
    pub in_degree: HashMap<(i32, i32), usize>,
    // --- Modify Undo/Redo State Storage ---
    #[cfg(feature = "undo_state")]
    undo_stack: Vec<PreviousCellState>, // Use a Vec for undo history [6, 7]
    #[cfg(feature = "undo_state")]
    redo_stack: Vec<PreviousCellState>, // Use a Vec for redo history [6, 7]
                                        // --- End Modifications ---
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
        }
    }
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
            // --- Initialize Undo/Redo Stacks ---
            #[cfg(feature = "undo_state")]
            undo_stack: Vec::with_capacity(MAX_UNDO_LEVELS), // Initialize empty stacks [6, 7]
            #[cfg(feature = "undo_state")]
            redo_stack: Vec::new(), // Redo stack often doesn't need strict capacity
                                    // --- End Initialization ---
        })
    }

    // --- Additions for Undo State ---
    // --- Helper to capture state (used by undo and redo) ---
    #[cfg(feature = "undo_state")] // <-- Update feature name
    pub fn capture_current_cell_state(&self, row: i32, col: i32) -> PreviousCellState {
        // This is essentially the same as capture_previous_state,
        // but the name clarifies its use in undo/redo logic.
        if let Some(cell) = self.cells.get(&(row, col)) {
            PreviousCellState {
                row,
                col,
                previous_formula_idx: cell.formula_idx,
                previous_value: cell.value,
                previous_status: cell.status.clone(),
                previous_dependencies: cell.dependencies.clone(),
                previous_dependents_of_cell: cell.dependents.clone(),
            }
        } else {
            // Cell doesn't exist, capture default state
            PreviousCellState {
                row,
                col,
                previous_formula_idx: None,
                previous_value: 0,
                previous_status: CellStatus::Ok,
                previous_dependencies: HashSet::new(),
                previous_dependents_of_cell: HashSet::new(),
            }
        }
    }

    // --- End Helper ---

    // Helper method to get or create a cell
    pub fn get_or_create_cell(&mut self, row: i32, col: i32) -> &mut Cell {
        if !self.cells.contains_key(&(row, col)) {
            self.cells.insert(
                (row, col),
                Cell {
                    value: 0,
                    formula_idx: None,
                    status: CellStatus::Ok,
                    dependencies: HashSet::new(),
                    dependents: HashSet::new(),
                    // Initialize cell history if feature is enabled
                    #[cfg(feature = "cell_history")]
                    history: VecDeque::with_capacity(MAX_HISTORY_SIZE),
                },
            );
        }
        self.cells.get_mut(&(row, col)).unwrap()
    }

    // Helper method to get cell value (returns 0 for non-existent cells)
    pub fn get_cell_value(&self, row: i32, col: i32) -> i32 {
        self.cells.get(&(row, col)).map_or(0, |cell| cell.value)
    }

    // Helper method to get cell status (returns Ok for non-existent cells)
    pub fn get_cell_status(&self, row: i32, col: i32) -> CellStatus {
        self.cells
            .get(&(row, col))
            .map_or(CellStatus::Ok, |cell| cell.status.clone())
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

    // Helper to update cell value and potentially its history
    pub fn update_cell_value(
        &mut self,
        row: i32,
        col: i32,
        new_value: i32,
        new_status: CellStatus,
    ) {
        let cell = self.get_or_create_cell(row, col);

        // --- Additions for Cell History ---
        // Only add to history if the value actually changes and feature is enabled
        #[cfg(feature = "cell_history")]
        {
            if cell.value != new_value {
                if cell.history.len() == MAX_HISTORY_SIZE {
                    cell.history.pop_front(); // Remove the oldest value
                }
                cell.history.push_back(cell.value); // Store the *current* value before overwriting
            }
        }
        // --- End Additions ---

        cell.value = new_value;
        cell.status = new_status;
    }
    // Add getter for cell history if feature enabled
    #[cfg(feature = "cell_history")]
    pub fn get_cell_history(&self, row: i32, col: i32) -> Option<Vec<i32>> {
        self.cells
            .get(&(row, col))
            .map(|cell| cell.history.iter().cloned().collect())
    }

    // Update cell formula (rewritten to use the sparse representation)
    pub fn update_cell_formula(
        &mut self,
        row: i32,
        col: i32,
        formula: &str,
        status_msg: &mut String,
    ) {
        // --- Additions for Undo State ---

        // // Clear the redo state whenever a new action is taken
        // #[cfg(feature = "undo_state")] // <-- Update feature name
        // {
        //     self.redo_state = None; // Any new edit invalidates the redo history
        // }
        // Capture state BEFORE any modification, only if feature is enabled
        #[cfg(feature = "undo_state")]
        let captured_prev_state = self.capture_current_cell_state(row, col);
        // --- End Additions ---

        if valid_formula(self, formula, status_msg) != 0 {
            status_msg.clear();
            status_msg.push_str("Unrecognized");
            return;
        }
        status_msg.clear();
        status_msg.push_str("Ok");

        #[cfg(feature = "undo_state")]
        {
            // Push the state *before* the change onto the undo stack
            self.undo_stack.push(captured_prev_state);

            // Enforce the history limit on the undo stack
            if self.undo_stack.len() > MAX_UNDO_LEVELS {
                self.undo_stack.remove(0); // Remove the oldest state [6, 7]
            }

            // Any new action clears the redo stack [7]
            self.redo_stack.clear();
        }

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
            if dep_row >= 0
                && dep_row < self.total_rows
                && dep_col >= 0
                && dep_col < self.total_cols
            {
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
            if dep_row >= 0
                && dep_row < self.total_rows
                && dep_col >= 0
                && dep_col < self.total_cols
            {
                // Store the dependency in the current cell
                self.get_or_create_cell(row, col)
                    .dependencies
                    .insert((dep_row, dep_col));

                // Store the current cell as a dependent of the dependency cell
                self.get_or_create_cell(dep_row, dep_col)
                    .dependents
                    .insert((row, col));
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
                let existing_idx = self
                    .formula_storage
                    .iter()
                    .position(|f| f == &old_formula_str);
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
                if dep_row >= 0
                    && dep_row < self.total_rows
                    && dep_col >= 0
                    && dep_col < self.total_cols
                {
                    // Add dependency to current cell
                    self.get_or_create_cell(row, col)
                        .dependencies
                        .insert((dep_row, dep_col));

                    // Add current cell as dependent to dependency cell
                    self.get_or_create_cell(dep_row, dep_col)
                        .dependents
                        .insert((row, col));
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
            crate::parser::evaluate_formula(
                &sheet_clone,
                formula,
                row,
                col,
                &mut error_flag,
                &mut s_msg,
            )
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
            // Set the value and status first
            {
                let cell = self.get_or_create_cell(row, col);
                #[cfg(feature = "cell_history")]
                {
                    if cell.value != new_val {
                        if cell.history.len() == 10 {
                            cell.history.pop_front(); // Remove the oldest value
                        }
                        cell.history.push_back(cell.value); // Store the *current* value before overwriting
                    }
                }
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
    // --- Apply a captured state (Helper for Undo/Redo) ---
    #[cfg(feature = "undo_state")] // <-- Update feature name
    pub fn apply_state(&mut self, state_to_apply: &PreviousCellState, status_msg: &mut String) {
        let row = state_to_apply.row;
        let col = state_to_apply.col;

        // 1. Get current dependencies before overwriting
        let current_deps = self
            .cells
            .get(&(row, col))
            .map_or(HashSet::new(), |c| c.dependencies.clone());

        // 2. Restore the cell's core properties
        {
            let cell = self.get_or_create_cell(row, col);
            cell.value = state_to_apply.previous_value;
            cell.status = state_to_apply.previous_status.clone();
            cell.formula_idx = state_to_apply.previous_formula_idx;
            cell.dependencies = state_to_apply.previous_dependencies.clone();
            cell.dependents = state_to_apply.previous_dependents_of_cell.clone();
        }

        // 3. Update dependent links based on the change
        // Remove the current cell from the dependents list of its *current* dependencies
        for &(dep_row, dep_col) in &current_deps {
            if let Some(dep_cell) = self.cells.get_mut(&(dep_row, dep_col)) {
                dep_cell.dependents.remove(&(row, col));
            }
        }
        // Add the current cell back to the dependents list of its *applied* dependencies
        for &(dep_row, dep_col) in &state_to_apply.previous_dependencies {
            self.get_or_create_cell(dep_row, dep_col)
                .dependents
                .insert((row, col));
        }

        // 4. Mark dirty and recalculate
        self.dirty_cells.insert((row, col));
        mark_cell_and_dependents_dirty(self, row, col);
        crate::parser::invalidate_cache_for_cell(row, col);
        recalc_affected(self, status_msg); // Recalculate using passed status_msg
    }
    // --- End Apply State Helper ---

    // --- Modify Undo Method for multi-level ---
    #[cfg(feature = "undo_state")]
    pub fn undo(&mut self, status_msg: &mut String) {
        status_msg.clear();

        // Pop from undo_stack if not empty [6, 7]
        if let Some(state_to_restore) = self.undo_stack.pop() {
            // Capture the current state *before* undoing, for REDO
            let state_before_undo =
                self.capture_current_cell_state(state_to_restore.row, state_to_restore.col);
            // Push the captured state onto the redo stack [6, 7]
            self.redo_stack.push(state_before_undo);
            // Note: Redo stack size limit isn't typically enforced strictly,
            // but could be added here if needed.

            // Apply the restored state using the helper
            self.apply_state(&state_to_restore, status_msg);

            if status_msg.is_empty() || status_msg == "Ok" {
                status_msg.clear();
                status_msg.push_str("Undo successful");
            }
        } else {
            status_msg.push_str("Nothing to undo");
        }
    }
    // --- End Undo Method ---
    // --- End Undo Method ---

    // --- Modify Redo Method for multi-level ---
    #[cfg(feature = "undo_state")]
    pub fn redo(&mut self, status_msg: &mut String) {
        status_msg.clear();

        // Pop from redo_stack if not empty [6, 7]
        if let Some(state_to_redo) = self.redo_stack.pop() {
            // Capture the state *before* redoing, for future UNDO
            let state_before_redo =
                self.capture_current_cell_state(state_to_redo.row, state_to_redo.col);
            // Push the captured state back onto the undo stack [6, 7]
            self.undo_stack.push(state_before_redo);
            // Enforce history limit on undo stack again after redo
            if self.undo_stack.len() > MAX_UNDO_LEVELS {
                self.undo_stack.remove(0);
            }

            // Apply the redone state using the helper
            self.apply_state(&state_to_redo, status_msg);

            if status_msg.is_empty() || status_msg == "Ok" {
                status_msg.clear();
                status_msg.push_str("Redo successful");
            }
        } else {
            status_msg.push_str("Nothing to redo");
        }
    }
    // --- End Redo Method ---
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
    if col_val == 0 {
        return None;
    }
    let col = col_val - 1;
    let mut row_val = 0;
    for ch in name[pos..].chars() {
        if ch.is_digit(10) {
            row_val = row_val * 10 + (ch as i32 - '0' as i32);
        } else {
            return None;
        }
    }
    if row_val <= 0 {
        return None;
    }
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

    if formula.starts_with("IF(") && cfg!(feature = "advanced_formulas") {
        // must have two commas and closing ')'
        let inner = &formula[3..formula.len().saturating_sub(1)];
        if inner.split(',').count() != 3 {
            status_msg.push_str("IF needs 3 args");
            return 1;
        }
        return 0;
    }
    if formula.starts_with("COUNTIF(") && cfg!(feature = "advanced_formulas") {
        let inner = &formula[8..formula.len().saturating_sub(1)];
        if inner.split(',').count() != 2 {
            status_msg.push_str("COUNTIF needs 2 args");
            return 1;
        }
        return 0;
    }
    if formula.starts_with("SUMIF(") && cfg!(feature = "advanced_formulas") {
        let inner = &formula[6..formula.len().saturating_sub(1)];
        if inner.split(',').count() != 3 {
            status_msg.push_str("SUMIF needs 3 args");
            return 1;
        }
        return 0;
    }
    if formula.starts_with("ROUND(") && cfg!(feature = "advanced_formulas") {
        let inner = &formula[6..formula.len().saturating_sub(1)];
        if inner.split(',').count() != 2 {
            status_msg.push_str("ROUND needs 2 args");
            return 1;
        }
        return 0;
    }

    if formula.starts_with("MAX(")
        || formula.starts_with("MIN(")
        || formula.starts_with("SUM(")
        || formula.starts_with("AVG(")
        || formula.starts_with("STDEV(")
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
        let inner = &formula[pos + 1..formula.len() - 1];
        let mut inner = inner.trim().to_string();
        if let Some(colon) = inner.find(':') {
            inner.replace_range(colon..colon + 1, ":");
            let cell1 = inner[..colon].to_string();
            let cell2 = inner[colon + 1..].to_string();
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
        let inner = &formula[6..formula.len() - 1];
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
            if ch.is_alphabetic() {
                break;
            }
            p = &p[ch.len_utf8()..];
        }
        if p.is_empty() {
            break;
        }

        let start = p;
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() {
                p = &p[ch.len_utf8()..];
            } else {
                break;
            }
        }
        while let Some(ch) = p.chars().next() {
            if ch.is_digit(10) {
                p = &p[ch.len_utf8()..];
            } else {
                break;
            }
        }

        if p.starts_with(':') {
            p = &p[1..];
            let range_start2 = p;
            while let Some(ch) = p.chars().next() {
                if ch.is_alphabetic() {
                    p = &p[ch.len_utf8()..];
                } else {
                    break;
                }
            }
            while let Some(ch) = p.chars().next() {
                if ch.is_digit(10) {
                    p = &p[ch.len_utf8()..];
                } else {
                    break;
                }
            }

            let len1 = start.find(':').unwrap_or(0);
            let cell_ref1 = &start[..len1];
            let cell_ref2 = &range_start2[..(range_start2.len() - p.len())];

            if let (Some((r1, c1)), Some((r2, c2))) = (
                cell_name_to_coords(cell_ref1),
                cell_name_to_coords(cell_ref2),
            ) {
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
    let mut ready_cells: Vec<(i32, i32)> = in_degree
        .iter()
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
                let new_val = crate::parser::evaluate_formula(
                    &sheet_clone,
                    &formula,
                    row,
                    col,
                    &mut error_flag,
                    &mut s_msg,
                );

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
                    #[cfg(feature = "cell_history")]
                    {
                        if cell.value != new_val {
                            if cell.history.len() == 10 {
                                cell.history.pop_front(); // Remove the oldest value
                            }
                            cell.history.push_back(cell.value); // Store the *current* value before overwriting
                        }
                    }
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
    let cells_with_cycles: Vec<(i32, i32)> = in_degree
        .iter()
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
pub fn build_dependency_graph(
    sheet: &Spreadsheet,
    row: i32,
    col: i32,
    dependencies_map: &mut HashMap<(i32, i32), HashSet<(i32, i32)>>,
    to_process: &mut HashSet<(i32, i32)>,
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
                dependencies_map
                    .entry((r, c))
                    .or_insert_with(HashSet::new)
                    .insert((dep_row, dep_col));

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
pub fn extract_dependencies_without_self(
    formula: &str,
    total_rows: i32,
    total_cols: i32,
) -> HashSet<(i32, i32)> {
    // For large range expressions, use a more space-efficient implementation
    // if formula.len() > 100 && formula.contains(':') {
    //     return extract_range_dependencies_optimized(formula, total_rows, total_cols);
    // }

    // Original implementation for smaller formulas
    let mut deps: HashSet<(i32, i32)> = HashSet::new();
    let mut p = formula;

    while !p.is_empty() {
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() {
                break;
            }
            p = &p[ch.len_utf8()..];
        }
        if p.is_empty() {
            break;
        }

        let start = p;
        while let Some(ch) = p.chars().next() {
            if ch.is_alphabetic() {
                p = &p[ch.len_utf8()..];
            } else {
                break;
            }
        }
        while let Some(ch) = p.chars().next() {
            if ch.is_digit(10) {
                p = &p[ch.len_utf8()..];
            } else {
                break;
            }
        }

        if p.starts_with(':') {
            p = &p[1..];
            let range_start2 = p;
            while let Some(ch) = p.chars().next() {
                if ch.is_alphabetic() {
                    p = &p[ch.len_utf8()..];
                } else {
                    break;
                }
            }
            while let Some(ch) = p.chars().next() {
                if ch.is_digit(10) {
                    p = &p[ch.len_utf8()..];
                } else {
                    break;
                }
            }

            let len1 = start.find(':').unwrap_or(0);
            let cell_ref1 = &start[..len1];
            let cell_ref2 = &range_start2[..(range_start2.len() - p.len())];

            if let (Some((r1, c1)), Some((r2, c2))) = (
                cell_name_to_coords(cell_ref1),
                cell_name_to_coords(cell_ref2),
            ) {
                if r1 >= 0
                    && r1 < total_rows
                    && c1 >= 0
                    && c1 < total_cols
                    && r2 >= 0
                    && r2 < total_rows
                    && c2 >= 0
                    && c2 < total_cols
                {
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
// pub fn extract_range_dependencies_optimized(formula: &str, total_rows: i32, total_cols: i32) -> HashSet<(i32, i32)> {
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
    sheet: &'a Spreadsheet,
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

#[cfg(test)]
mod tests {
    // everything declared in sheet.rs
    use super::*;
    use crate::parser::clear_range_cache;
    use crate::parser::{
        // for the cache/history tests
        evaluate_formula,
    };

    // only compile these if the cli_app feature is on
    #[cfg(feature = "cli_app")]
    use super::*; // now brings in sheet.rs items via the outer super
                  // import your CLI helpers:
    use crate::cli_app::{clamp_viewport_hz, clamp_viewport_ve, col_to_letters, process_command};

    #[test]
    fn cell_name_roundtrip_and_bounds() {
        assert_eq!(cell_name_to_coords("AA10"), Some((9, 26)));
        assert_eq!(coords_to_cell_name(9, 26), "AA10");
        // invalid
        assert!(cell_name_to_coords("1A").is_none());
    }

    #[test]
    fn valid_formula_simple_and_errors() {
        let sheet = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        // plain number or ref
        assert_eq!(valid_formula(&sheet, "42", &mut msg), 0);
        assert_eq!(valid_formula(&sheet, "A1", &mut msg), 0);
        // bad formula
        assert_eq!(valid_formula(&sheet, "", &mut msg), 1);
        assert!(msg.contains("Empty"));
    }

    #[test]
    fn extract_dependencies_and_circular() {
        let deps = extract_dependencies(&Spreadsheet::new(5, 5), "B2:C3");
        assert!(deps.contains(&(1, 1)) && deps.contains(&(2, 2)));
        // manual cycle
        let mut s = Spreadsheet::new(2, 2);
        s.get_or_create_cell(0, 0).dependencies.insert((0, 1));
        s.get_or_create_cell(0, 1).dependencies.insert((0, 0));
        assert!(has_circular_dependency(&s, 0, 0));
        assert!(has_circular_dependency_by_index(&s, 0, 0));
    }

    #[test]
    fn clear_and_invalidate_range_cache() {
        let mut s = Spreadsheet::new(2, 2);
        s.update_cell_value(0, 0, 3, CellStatus::Ok);
        s.update_cell_value(0, 1, 5, CellStatus::Ok);
        let mut msg = String::new();
        let cs = CloneableSheet::new(&s);
        assert_eq!(
            crate::parser::evaluate_formula(&cs, "SUM(A1:B1)", 0, 0, &mut 0, &mut msg),
            8
        );
        clear_range_cache();
        // change value and re-eval
        s.update_cell_value(0, 0, 7, CellStatus::Ok);
        let cs2 = CloneableSheet::new(&s);
        assert_eq!(
            crate::parser::evaluate_formula(&cs2, "SUM(A1:B1)", 0, 0, &mut 0, &mut msg),
            12
        );
    }

    fn recalc_dirty_and_error_propagation() {
        // Build a little chain of dependents A1→A2→A3→A4
        let mut sheet = Spreadsheet::new(4, 1);
        sheet.get_or_create_cell(0, 0).dependents.insert((1, 0));
        sheet.get_or_create_cell(1, 0).dependents.insert((2, 0));
        sheet.get_or_create_cell(2, 0).dependents.insert((3, 0));

        // Initially clean:
        assert!(sheet.dirty_cells.is_empty());

        // Mark A1 dirty and propagate
        mark_cell_and_dependents_dirty(&mut sheet, 0, 0);

        // A2 and A3 get marked; A4 does _not_
        assert!(sheet.dirty_cells.contains(&(1, 0)), "A2 should be dirty");
        assert!(sheet.dirty_cells.contains(&(2, 0)), "A3 should be dirty");
        assert!(
            !sheet.dirty_cells.contains(&(3, 0)),
            "A4 should NOT be dirty"
        );
    }

    #[test]
    fn mark_and_dirty_cells() {
        let mut s = Spreadsheet::new(4, 1);
        // manually wire dependents: A1→A2→A3→A4
        for i in 0..3 {
            s.get_or_create_cell(i, 0).dependents.insert((i + 1, 0));
        }
        mark_cell_and_dependents_as_error(&mut s, 0, 0);
        for i in 0..4 {
            assert_eq!(s.get_cell_status(i, 0), CellStatus::Error);
        }
    }

    use super::*; // brings Spreadsheet, CellStatus, etc.

    #[test]
    fn get_or_create_cell_tracks_dependencies_and_dependents() {
        let mut sheet = Spreadsheet::new(3, 3);
        // make A1 depend on B2
        sheet.get_or_create_cell(0, 0).dependencies.insert((1, 1));
        sheet.get_or_create_cell(1, 1).dependents.insert((0, 0));
        // verify detection
        assert!(has_circular_dependency_by_index(&sheet, 0, 0) == false);
    }

    #[test]
    fn mark_error_propagates_through_chain() {
        let mut sheet = Spreadsheet::new(3, 1);
        // A1→A2→A3
        sheet.get_or_create_cell(0, 0).dependents.insert((1, 0));
        sheet.get_or_create_cell(1, 0).dependents.insert((2, 0));
        mark_cell_and_dependents_as_error(&mut sheet, 0, 0);
        assert_eq!(sheet.get_cell_status(2, 0), CellStatus::Error);
    }

    fn name_and_coords_roundtrip() {
        assert_eq!(cell_name_to_coords("A1"), Some((0, 0)));
        assert_eq!(cell_name_to_coords("AA10"), Some((9, 26)));
        assert_eq!(coords_to_cell_name(0, 0), "A1");
        assert_eq!(coords_to_cell_name(9, 26), "AA10");
    }

    #[test]
    fn valid_formula_basic_cases() {
        let sheet = Spreadsheet::new(3, 3);
        let mut msg = String::new();

        // empty
        assert_eq!(valid_formula(&sheet, "", &mut msg), 1);
        assert_eq!(msg, "Empty formula");

        // plain number
        msg.clear();
        assert_eq!(valid_formula(&sheet, "123", &mut msg), 0);

        // OOB cell
        msg.clear();
        assert_eq!(valid_formula(&sheet, "Z99", &mut msg), 1);
        assert_eq!(msg, "Cell reference out of bounds");
    }

    #[test]
    fn extract_dependencies_single_and_range() {
        let sheet = Spreadsheet::new(2, 2);

        // single
        let single = extract_dependencies(&sheet, "B2");
        let mut want = std::collections::HashSet::new();
        want.insert((1, 1));
        assert_eq!(single, want);

        // range
        let range = extract_dependencies(&sheet, "SUM(A1:B2)");
        let mut expected = std::collections::HashSet::new();
        expected.insert((0, 0));
        expected.insert((0, 1));
        expected.insert((1, 0));
        expected.insert((1, 1));
        assert_eq!(range, expected);
    }

    #[test]
    fn update_formula_chain_propagates_immediately() {
        let mut sheet = Spreadsheet::new(4, 1);
        let mut status = String::new();

        // Build: A1=2, A2=A1*2, A3=A2*2, A4=A3*2
        sheet.update_cell_formula(0, 0, "2", &mut status);
        sheet.update_cell_formula(1, 0, "A1*2", &mut status);
        sheet.update_cell_formula(2, 0, "A2*2", &mut status);
        sheet.update_cell_formula(3, 0, "A3*2", &mut status);

        // Now change A1 to 3
        sheet.update_cell_formula(0, 0, "3", &mut status);

        // All downstream cells should have updated right away:
        assert_eq!(sheet.get_cell_value(0, 0), 3);
        assert_eq!(sheet.get_cell_value(1, 0), 6);
        assert_eq!(sheet.get_cell_value(2, 0), 12);
        assert_eq!(sheet.get_cell_value(3, 0), 24);
    }
    #[test]
    fn raw_content_and_defaults() {
        let mut sheet = Spreadsheet::new(2, 2);
        // no formula yet
        assert_eq!(sheet.get_cell_raw_content(0, 0), "");
        // default value/status
        assert_eq!(sheet.get_cell_value(0, 0), 0);
        assert_eq!(sheet.get_cell_status(0, 0), CellStatus::Ok);
    }

    // …and tests for undo/redo, history (if cell_history feature enabled), etc.
    // at the bottom of src/sheet.rs

    use super::*; // brings in Spreadsheet, CellStatus, valid_formula, extract_dependencies, etc.
    use std::collections::HashSet;

    #[test]
    fn new_and_name_roundtrip() {
        let s = Spreadsheet::new(3, 4);
        assert_eq!(s.total_rows, 3);
        assert_eq!(s.total_cols, 4);
        assert!(s.cells.is_empty());
        // name conversions
        assert_eq!(cell_name_to_coords("B2"), Some((1, 1)));
        assert_eq!(coords_to_cell_name(1, 1), "B2");
    }

    #[test]
    fn set_value_and_raw_content() {
        let mut s = Spreadsheet::new(2, 2);
        s.update_cell_value(0, 0, 99, CellStatus::Ok);
        assert_eq!(s.get_cell_value(0, 0), 99);
        assert_eq!(s.get_cell_raw_content(0, 0), "");
    }

    #[test]
    fn set_formula_and_recalc_chain() {
        let mut s = Spreadsheet::new(2, 2);
        let mut status = String::new();
        s.update_cell_formula(0, 0, "5", &mut status);
        s.update_cell_formula(1, 0, "A1*2", &mut status);
        assert_eq!(s.get_cell_value(1, 0), 10);
    }

    #[test]
    fn recalc_affected_propagates_and_clears_dirty() {
        let mut s = Spreadsheet::new(3, 1);
        let mut status = String::new();
        // A1=1, A2=A1+1, A3=A2+1
        s.update_cell_formula(0, 0, "1", &mut status);
        s.update_cell_formula(1, 0, "A1+1", &mut status);
        s.update_cell_formula(2, 0, "A2+1", &mut status);

        // dirty just A1
        s.dirty_cells.insert((0, 0));
        recalc_affected(&mut s, &mut status);

        assert!(s.dirty_cells.is_empty());
        assert_eq!(s.get_cell_value(2, 0), 3);
    }

    #[test]
    fn extract_and_validate() {
        let s = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        // valid literal formula
        assert_eq!(valid_formula(&s, "123", &mut msg), 0);
        // invalid cell ref
        assert_eq!(valid_formula(&s, "X9", &mut msg), 1);

        // extract full 3×3 range
        let deps = extract_dependencies(&s, "A1:C3");
        let mut want = HashSet::new();
        for r in 0..3 {
            for c in 0..3 {
                want.insert((r, c));
            }
        }
        assert_eq!(deps, want);
    }

    #[test]
    fn error_propagation_on_div_zero() {
        let mut s = Spreadsheet::new(2, 2);
        let mut msg = String::new();
        s.update_cell_formula(0, 0, "10", &mut msg);
        s.update_cell_formula(0, 1, "0", &mut msg);
        s.update_cell_formula(1, 0, "A1/B1", &mut msg);
        assert_eq!(s.get_cell_status(1, 0), CellStatus::Error);
    }

    #[test]
    fn test_col_to_letters_cli() {
        assert_eq!(crate::cli_app::col_to_letters(0), "A");
        assert_eq!(crate::cli_app::col_to_letters(25), "Z");
        assert_eq!(crate::cli_app::col_to_letters(26), "AA");
        assert_eq!(crate::cli_app::col_to_letters(701), "ZZ");
        assert_eq!(crate::cli_app::col_to_letters(702), "AAA");
    }

    #[test]
    fn test_clamp_viewport() {
        let mut r = 50;
        clamp_viewport_ve(40, &mut r);
        assert_eq!(r, 40, "50 > 40, so we do 50 - 10 = 40");

        let mut r2 = -5;
        clamp_viewport_ve(100, &mut r2);
        assert_eq!(r2, 0, "-5 < 0, clamps up to 0");

        let mut c = 95;
        clamp_viewport_hz(90, &mut c);
        assert_eq!(c, 85, "95 > 90, so we do 95 - 10 = 85");

        let mut c2 = -1;
        clamp_viewport_hz(10, &mut c2);
        assert_eq!(c2, 0, "-1 < 0, clamps up to 0");
    }
    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
    // 2) process_command branches
    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––

    #[test]
    fn test_process_command_scroll_and_invalid() {
        let mut sheet = Spreadsheet::new(5, 5);
        let mut status = String::new();

        process_command(&mut sheet, "scroll_to B2", &mut status);
        assert_eq!((sheet.top_row, sheet.left_col), (1, 1));

        process_command(&mut sheet, "scroll_to Z9", &mut status);
        assert!(status.contains("out of bounds"));

        process_command(&mut sheet, "foo", &mut status);
        assert_eq!(status, "unrecognized cmd");
    }

    #[test]
    fn test_clear_cache_and_history_without_feature() {
        let mut sheet = Spreadsheet::new(3, 3);
        let mut status = String::new();

        // seed the cache
        sheet.cache.insert(
            "X".into(),
            CachedRange {
                value: 1,
                dependencies: Default::default(),
            },
        );

        process_command(&mut sheet, "clear_cache", &mut status);
        assert_eq!(status, "Cache cleared");
        assert!(sheet.cache.is_empty());

        // history should say "not enabled"
        process_command(&mut sheet, "history A1", &mut status);
        assert_eq!(status, "Cell history feature is not enabled.");
    }

    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
    // 3) Undo/Redo & History (if features enabled)
    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––

    #[cfg(feature = "undo_state")]
    #[test]
    fn test_undo_redo_basic() {
        let mut sheet = Spreadsheet::new(2, 2);
        let mut status = String::new();
        sheet.update_cell_formula(0, 0, "5", &mut status);
        sheet.update_cell_formula(0, 0, "6", &mut status);
        super::cli_app::process_command(&mut sheet, "undo", &mut status);
        assert_eq!(sheet.get_cell_value(0, 0), 5);
        super::cli_app::process_command(&mut sheet, "redo", &mut status);
        assert_eq!(sheet.get_cell_value(0, 0), 6);
    }

    #[cfg(feature = "cell_history")]
    #[test]
    fn test_cell_history_feature() {
        let mut sheet = Spreadsheet::new(2, 2);
        let mut status = String::new();
        sheet.update_cell_formula(0, 0, "5", &mut status);
        sheet.update_cell_formula(0, 0, "7", &mut status);
        super::cli_app::process_command(&mut sheet, "history A1", &mut status);
        assert_eq!(status, "History displayed");
    }

    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
    // 4) mark_dirty & recalc + dependency graph
    //––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––

    #[test]
    fn test_dirty_and_recalc_chain() {
        let mut sheet = Spreadsheet::new(4, 1);
        sheet.get_or_create_cell(0, 0).dependents.insert((1, 0));
        sheet.get_or_create_cell(1, 0).dependents.insert((2, 0));
        sheet.get_or_create_cell(2, 0).dependents.insert((3, 0));

        // mark dirty
        mark_cell_and_dependents_dirty(&mut sheet, 0, 0);
        assert!(sheet.dirty_cells.contains(&(1, 0)));
        assert!(sheet.dirty_cells.contains(&(2, 0)));
        assert!(!sheet.dirty_cells.contains(&(3, 0)));

        // recalc with a simple chain of formulas
        let mut status = String::new();
        sheet.update_cell_formula(0, 0, "1", &mut status);
        sheet.update_cell_formula(1, 0, "A1+1", &mut status);
        sheet.update_cell_formula(2, 0, "A2+1", &mut status);
        sheet.dirty_cells.insert((0, 0));
        recalc_affected(&mut sheet, &mut status);
        assert_eq!(sheet.get_cell_value(2, 0), 3);
        assert!(sheet.dirty_cells.is_empty());
    }

    use super::*;
    // 2) pull in your CLI layer so `process_command` is visible:
    use crate::cli_app;

    // ────────────────────────────────────────────────────────────────────────────────
    // expose private parser functions & types to our #[cfg(test)] module
    #[cfg(test)]
    pub(crate) mod test_exports {
        pub use crate::parser::{
            clear_range_cache, evaluate_ast, evaluate_large_range, evaluate_range_function,
            invalidate_cache_for_cell, parse_expr, parse_factor, parse_term, ASTNode, RANGE_CACHE,
        };
    }

    #[cfg(test)]
    mod parser_internal_tests {
        use super::test_exports::*;
        use crate::sheet::{CellStatus, CloneableSheet, Spreadsheet};

        #[test]
        fn test_skip_spaces_and_number_parsing() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut s = "   -123abc";
            let mut err = 0;
            let v = parse_factor(&cs, &mut s, 0, 0, &mut err);
            assert_eq!(v, -123);
            assert_eq!(err, 0);
        }

        #[test]
        fn test_parse_comparisons_and_binary() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;

            let mut s1 = "2>1";
            assert_eq!(parse_expr(&cs, &mut s1, 0, 0, &mut err), 1);

            let mut s2 = "5<=3";
            assert_eq!(parse_expr(&cs, &mut s2, 0, 0, &mut err), 0);

            let mut s3 = "4==4";
            assert_eq!(parse_expr(&cs, &mut s3, 0, 0, &mut err), 1);

            let mut s4 = "3+2*4-5/5";
            assert_eq!(parse_expr(&cs, &mut s4, 0, 0, &mut err), 10);
        }

        #[test]
        fn test_evaluate_range_function_errors_and_success() {
            let mut sheet = Spreadsheet::new(2, 2);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;

            // syntax error
            assert_eq!(evaluate_range_function(&cs, "SUM", "A1B2", &mut err), 0);
            assert_eq!(err, 1);

            // out‐of‐bounds
            err = 0;
            assert_eq!(evaluate_range_function(&cs, "SUM", "A1:C1", &mut err), 0);
            assert_eq!(err, 4);

            // valid
            sheet.update_cell_value(0, 0, 1, CellStatus::Ok);
            sheet.update_cell_value(0, 1, 2, CellStatus::Ok);
            let cs2 = CloneableSheet::new(&*sheet);
            err = 0;
            assert_eq!(evaluate_range_function(&cs2, "SUM", "A1:B1", &mut err), 3);
        }

        #[test]
        fn test_clear_and_invalidate_cache_for_cell() {
            clear_range_cache();
            // manually inject into the RANGE_CACHE:
            RANGE_CACHE.with(|c| {
                c.borrow_mut()
                    .insert("X".into(), (5, [(0, 0)].iter().cloned().collect()));
            });
            invalidate_cache_for_cell(0, 0);
            RANGE_CACHE.with(|c| assert!(c.borrow().is_empty()));
        }

        #[test]
        fn test_evaluate_ast_literals_and_sleep() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;

            let lit = ASTNode::Literal(7);
            assert_eq!(evaluate_ast(&cs, &lit, 0, 0, &mut err), 7);

            let sf = ASTNode::SleepFunction(Box::new(ASTNode::Literal(-2)));
            assert_eq!(evaluate_ast(&cs, &sf, 0, 0, &mut err), -2);
        }
    }

    #[test]
    fn test_extract_dependencies_without_self_simple_range() {
        // A1:B2 should give (0,0),(0,1),(1,0),(1,1)
        let deps = extract_dependencies_without_self("A1:B2", 10, 10);
        let mut want = std::collections::HashSet::new();
        want.insert((0, 0));
        want.insert((0, 1));
        want.insert((1, 0));
        want.insert((1, 1));
        assert_eq!(deps, want);
    }

    #[test]
    fn test_extract_dependencies_without_self_reversed_range() {
        // Even if the bounds are reversed, it should normalize
        let deps = extract_dependencies_without_self("B2:A1", 10, 10);
        let mut want = std::collections::HashSet::new();
        want.insert((0, 0));
        want.insert((0, 1));
        want.insert((1, 0));
        want.insert((1, 1));
        assert_eq!(deps, want);
    }

    #[test]
    fn test_extract_dependencies_without_self_multi_letter_cols() {
        // AA1:AB2 maps to rows 0–1, cols 26→27 (0‐based 26 and 27)
        let deps = extract_dependencies_without_self("AA1:AB2", 100, 100);
        let mut want = std::collections::HashSet::new();
        for r in 0..=1 {
            for c in 26..=27 {
                want.insert((r, c));
            }
        }
        assert_eq!(deps, want);
    }

    #[test]
    fn test_extract_dependencies_without_self_single_cell() {
        // Single cell reference
        let deps = extract_dependencies_without_self("C3", 5, 5);
        let mut want = std::collections::HashSet::new();
        want.insert((2, 2));
        assert_eq!(deps, want);
    }

    #[test]
    fn test_extract_dependencies_without_self_out_of_bounds() {
        // Z1 is col 25, but total_cols=5 → out of bounds, so skip
        let deps = extract_dependencies_without_self("Z1:Z1", 5, 5);
        assert!(deps.is_empty());
    }

    // Add to the `#[cfg(test)] mod tests { ... }` in src/sheet.rs

    #[test]
    fn valid_formula_range_errors() {
        let mut sheet = Spreadsheet::new(5, 5);
        let mut msg = String::new();

        // Missing colon in range
        assert_eq!(valid_formula(&sheet, "SUM(A1A2)", &mut msg), 1);
        assert!(msg.contains("Missing colon in range"));

        // Second cell reference out of bounds
        msg.clear();
        assert_eq!(valid_formula(&sheet, "SUM(A1:Z10)", &mut msg), 1);
        assert!(msg.contains("Second cell reference out of bounds"));

        // Invalid range order
        msg.clear();
        assert_eq!(valid_formula(&sheet, "SUM(B2:A1)", &mut msg), 1);
        assert!(msg.contains("Invalid range order"));
    }

    fn valid_formula_operator_not_found() {
        let sheet = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        let code = valid_formula(&sheet, "XYZ", &mut msg);
        assert_eq!(code, 1);
        assert_eq!(msg, "Operator not found");
    }

    #[test]
    fn valid_formula_sleep_variants() {
        let sheet = Spreadsheet::new(2, 2);
        let mut msg = String::new();
        // Missing closing parenthesis
        assert_eq!(valid_formula(&sheet, "SLEEP(1", &mut msg), 1);
        assert_eq!(msg, "Missing closing parenthesis in SLEEP");
        // Integer argument
        msg.clear();
        let code_int = valid_formula(&sheet, "SLEEP(5)", &mut msg);
        assert_eq!(code_int, 0);
        assert_eq!(msg, "");
        // Valid cell reference
        msg.clear();
        let code_ref = valid_formula(&sheet, "SLEEP(A1)", &mut msg);
        assert_eq!(code_ref, 0);
        // Out-of-bounds cell reference
        msg.clear();
        let code_oob = valid_formula(&sheet, "SLEEP(Z9)", &mut msg);
        assert_eq!(code_oob, 1);
        assert_eq!(msg, "Cell reference in up out of bounds");
    }

    fn recalc_detects_div_zero_and_marks_error() {
        let mut sheet = Spreadsheet::new(2, 1);
        let mut status = String::new();
        // A1 = 2
        sheet.update_cell_formula(0, 0, "2", &mut status);
        // A2 = A1/0 → division by zero
        sheet.update_cell_formula(1, 0, "A1/0", &mut status);
        sheet.dirty_cells.insert((1, 0));
        recalc_affected(&mut sheet, &mut status);
        assert_eq!(sheet.get_cell_status(1, 0), CellStatus::Error);
        assert_eq!(sheet.get_cell_value(1, 0), 0);
    }

    fn inject_formula(sheet: &mut Spreadsheet, formula: &'static str) {
        // Push into storage and point A1 at it
        let idx = sheet.formula_storage.len();
        sheet.formula_storage.push(formula.into());
        let cell = sheet.get_or_create_cell(0, 0);
        cell.formula_idx = Some(idx);
        // clear any old dependencies & dependents
        cell.dependencies.clear();
        cell.dependents.clear();
        sheet.dirty_cells.insert((0, 0));
    }
    #[test]
    fn recalc_detects_reversed_range_as_invalid_range() {
        let mut sheet = Spreadsheet::new(1, 1);
        let mut status = String::new();

        // SUM(A2:A1) is syntactically valid, but 2 > 1 should trigger error_flag=2
        inject_formula(&mut sheet, "SUM(A2:A1)");
        status.clear();
        recalc_affected(&mut sheet, &mut status);
        assert_eq!(status, "Invalid range");
    }

    #[test]
    fn recalc_detects_syntax_error_as_general_error() {
        let mut sheet = Spreadsheet::new(1, 1);
        let mut status = String::new();

        // "1?1" passes our simple valid_formula (it has an operator at index 1),
        // but parse_expr will set error_flag=1 on the '?'.
        inject_formula(&mut sheet, "1?1");
        status.clear();
        recalc_affected(&mut sheet, &mut status);
        assert_eq!(status, "Error in formula");
    }

    fn cloneable_sheet_get_cell_various() {
        let mut sheet = Spreadsheet::new(2, 2);
        // Set a cell to a non-default value and status
        sheet.update_cell_value(1, 1, 42, CellStatus::Error);
        let cs = CloneableSheet::new(&sheet);

        // Out-of-bounds coordinates should return None
        assert!(cs.get_cell(-1, 0).is_none());
        assert!(cs.get_cell(0, -1).is_none());
        assert!(cs.get_cell(2, 0).is_none());
        assert!(cs.get_cell(0, 2).is_none());

        // In-bounds but non-existent cell should return default view
        let view = cs.get_cell(0, 1).unwrap();
        assert_eq!(view.value, 0);
        assert_eq!(view.status, CellStatus::Ok);

        // Existing cell should return its actual value and status
        let view2 = cs.get_cell(1, 1).unwrap();
        assert_eq!(view2.value, 42);
        assert_eq!(view2.status, CellStatus::Error);
    }

    #[test]
    fn valid_formula_sleep_and_operator_and_format() {
        let sheet = Spreadsheet::new(3, 3);
        let mut msg = String::new();

        // Missing closing parenthesis in SLEEP
        let code = valid_formula(&sheet, "SLEEP(5", &mut msg);
        assert_eq!(code, 1);
        assert_eq!(msg, "Missing closing parenthesis in SLEEP");

        // Valid numeric SLEEP
        msg.clear();
        let code = valid_formula(&sheet, "SLEEP( 7 )", &mut msg);
        assert_eq!(code, 0);

        // Invalid cell reference in SLEEP
        msg.clear();
        let code = valid_formula(&sheet, "SLEEP(ABC)", &mut msg);
        assert_eq!(code, 1);
        assert_eq!(msg, "Invalid cell reference in SLEEP");

        // Operator not found
        msg.clear();
        let code = valid_formula(&sheet, "foobar", &mut msg);
        assert_eq!(code, 1);
        assert_eq!(msg, "Operator not found");

        // Invalid formula format (right side not int or cell)
        msg.clear();
        let code = valid_formula(&sheet, "A1+foo", &mut msg);
        assert_eq!(code, 1);
        assert_eq!(msg, "Invalid formula format");
    }

    fn process_command_scroll_to_invalid_cell() {
        let mut sheet = Spreadsheet::new(5, 5);
        let mut msg = String::new();
        process_command(&mut sheet, "scroll_to foo", &mut msg);
        assert!(
            msg.contains("Invalid cell"),
            "got `{}` expected substring `Invalid cell`",
            msg
        );
    }

    #[test]
    fn process_command_clear_cache() {
        let mut sheet = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        sheet.cache.insert(
            "X".into(),
            CachedRange {
                value: 1,
                dependencies: HashSet::new(),
            },
        );
        process_command(&mut sheet, "clear_cache", &mut msg);
        assert_eq!(msg, "Cache cleared");
        assert!(sheet.cache.is_empty());
    }

    #[test]
    fn cloneable_get_cell_bounds_and_default() {
        let sheet = Spreadsheet::new(2, 2);
        let cs = crate::sheet::CloneableSheet::new(&*sheet);
        // out of bounds → None
        assert!(cs.get_cell(-1, 0).is_none());
        assert!(cs.get_cell(0, -1).is_none());
        assert!(cs.get_cell(2, 0).is_none());
        // in bounds but never set → Some(default)
        let cv = cs.get_cell(1, 1).unwrap();
        assert_eq!(cv.value, 0);
        assert_eq!(cv.status, CellStatus::Ok);
    }

    #[test]
    fn extract_dependencies_without_self_multi_letter_cols() {
        let deps = crate::sheet::extract_dependencies_without_self("AA1:AB2", 100, 100);
        let mut want = HashSet::new();
        // AA → col 26, AB → col 27 (0-based)
        for r in 0..=1 {
            for c in 26..=27 {
                want.insert((r, c));
            }
        }
        assert_eq!(deps, want);
    }

    /// valid_formula: missing colon in range
    #[test]
    fn sheet_valid_formula_missing_colon() {
        let s = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        assert_eq!(valid_formula(&s, "SUM(A1A2)", &mut msg), 1);
        assert!(msg.contains("Missing colon"));
    }

    /// valid_formula: invalid range order
    #[test]
    fn sheet_valid_formula_invalid_range_order() {
        let s = Spreadsheet::new(3, 3);
        let mut msg = String::new();
        assert_eq!(valid_formula(&s, "SUM(B2:A1)", &mut msg), 1);
        assert!(msg.contains("Invalid range order"));
    }

    /// valid_formula: missing ')' on SLEEP
    #[test]
    fn sheet_valid_formula_sleep_missing_paren() {
        let s = Spreadsheet::new(1, 1);
        let mut msg = String::new();
        assert_eq!(valid_formula(&s, "SLEEP(1", &mut msg), 1);
        assert!(msg.contains("Missing closing parenthesis"));
    }

    /// valid_formula: invalid cell in SLEEP
    #[test]
    fn sheet_valid_formula_sleep_invalid_cell() {
        let s = Spreadsheet::new(1, 1);
        let mut msg = String::new();
        assert_eq!(valid_formula(&s, "SLEEP(B2)", &mut msg), 1);
        assert!(msg.contains("out of bounds"));
    }

    /// valid_formula: operator not found
    #[test]
    fn sheet_valid_formula_operator_not_found() {
        let s = Spreadsheet::new(1, 1);
        let mut msg = String::new();
        assert_eq!(valid_formula(&s, "foo", &mut msg), 1);
        assert!(msg.contains("Operator not found"));
    }

    /// valid_formula: invalid format

    /// recalc_affected: invalid range sets "Invalid range"

    /// recalc_affected: general error sets "Error in formula"

    /// extract_dependencies_without_self: two‐letter columns
    #[test]
    fn sheet_extract_deps_without_self_two_letter() {
        let deps = extract_dependencies_without_self("AA1:AB2", 100, 100);
        let mut want = HashSet::new();
        // AA→col 26, AB→col 27 (0-based)
        for r in 0..=1 {
            for c in 26..=27 {
                want.insert((r, c));
            }
        }
        assert_eq!(deps, want);
    }

    /// extract_dependencies: simple single + range
    #[test]
    fn sheet_extract_dependencies_basic() {
        let s = Spreadsheet::new(3, 3);
        let single = extract_dependencies(&s, "B2");
        assert_eq!(single, vec![(1, 1)].into_iter().collect());

        let range = extract_dependencies(&s, "SUM(A1:C1)");
        let mut want = HashSet::new();
        for c in 0..3 {
            want.insert((0, c));
        }
        assert_eq!(range, want);
    }
}
