#![allow(warnings)]

use spreadsheet::parser;
use spreadsheet::sheet;

#[cfg(feature = "cli_app")]
pub mod cli_app {
    // Use crate's modules
    use crate::parser::*;
    use crate::sheet::*;
    use std::env;
    use std::io::{self, Write};
    use std::time::Duration;
    use std::time::Instant;

    // Global variable as in C (check)
    static mut CHECK: bool = false;

    // Converts a 0-indexed column number into its corresponding letter string.
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

    // Clamps vertical viewport.
    fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
        if *start_row > total_rows {
            *start_row -= 10;
        } else if *start_row > (total_rows - 10) {
            *start_row = total_rows - 10;
        } else if *start_row < 0 {
            *start_row = 0;
        }
    }

    // Clamps horizontal viewport.
    fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
        if *start_col > total_cols {
            *start_col -= 10;
        } else if *start_col > (total_cols - 10) {
            *start_col = total_cols - 10;
        } else if *start_col < 0 {
            *start_col = 0;
        }
    }

    // Displays the grid (viewport 10x10).
    fn display_grid(sheet: &Spreadsheet) {
        let start_row = sheet.top_row;
        let start_col = sheet.left_col;
        let mut end_row = start_row + 10;
        let mut end_col = start_col + 10;
        if end_row > sheet.total_rows {
            end_row = sheet.total_rows;
        }
        if end_col > sheet.total_cols {
            end_col = sheet.total_cols;
        }

        // Print column headers.
        print!("     ");
        for c in start_col..end_col {
            let col_buf = col_to_letters(c);
            print!("{:<12}", col_buf);
        }
        println!();

        // Print rows with values.
        for r in start_row..end_row {
            print!("{:<4} ", r + 1);
            for c in start_col..end_col {
                // Get cell value from the sparse representation
                let status = sheet.get_cell_status(r, c);
                if status == CellStatus::Error {
                    print!("{:<12}", "ERR");
                } else {
                    print!("{:<12}", sheet.get_cell_value(r, c));
                }
            }
            println!();
        }
    }

    // Displays grid from a specified start.
    fn display_grid_from(sheet: &Spreadsheet, start_row: i32, start_col: i32) {
        // Calculate max displayable rows/columns
        let mut max_col = start_col + 10;
        if max_col > sheet.total_cols {
            max_col = sheet.total_cols;
        }

        let mut max_row = start_row + 10;
        if max_row > sheet.total_rows {
            max_row = sheet.total_rows;
        }

        // Always print at least column headers
        print!("     ");
        for c in start_col..max_col {
            let col_buf = col_to_letters(c);
            print!("{:<12}", col_buf);
        }
        println!();

        // Print rows with boundary checking
        for r in start_row..max_row {
            if r < 0 || r >= sheet.total_rows {
                continue;
            }

            print!("{:<4} ", r + 1);
            for c in start_col..max_col {
                if c < 0 || c >= sheet.total_cols {
                    print!("{:<12}", "--");
                    continue;
                }

                // Get cell value from the sparse representation
                let status = sheet.get_cell_status(r, c);
                if status == CellStatus::Error {
                    print!("{:<12}", "ERR");
                } else {
                    print!("{:<12}", sheet.get_cell_value(r, c));
                }
            }
            println!();
        }
    }

    // Process commands: scrolling, cell assignment, output control.
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
            // Clear both sheet cache and parser cache
            sheet.cache.clear();
            sheet.dirty_cells.clear();
            clear_range_cache();
            *status_msg = "Cache cleared".to_string();
        } else if cmd.starts_with("history") {
            let parts: Vec<&str> = cmd.split_whitespace().collect();
            if parts.len() == 2 {
                let cell_ref = parts[1];
                if let Some((row, col)) = cell_name_to_coords(cell_ref) {
                    // [1, 3]
                    if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                        *status_msg =
                            format!("Cell {} out of bounds", cell_ref.to_uppercase()).to_string();
                        //  status_msg.push_str(&format!("Cell {} out of bounds", cell_ref.to_uppercase()));
                    } else {
                        // --- Feature Check ---
                        #[cfg(feature = "cell_history")] // [4]
                        {
                            // Assuming get_cell_history exists in Spreadsheet impl [1]
                            match sheet.get_cell_history(row, col) {
                                Some(history) if !history.is_empty() => {
                                    // Print directly instead of using status_msg due to potential length
                                    println!("History for {}:", cell_ref.to_uppercase());
                                    // Print oldest first (index 1)
                                    for (i, val) in history.iter().enumerate() {
                                        println!("  {}: {}", i + 1, val);
                                    }
                                    let current_val = sheet.get_cell_value(row, col);
                                    println!("  Current: {}", current_val);
                                    *status_msg = "History displayed".to_string();
                                    // status_msg.push_str("History displayed"); // Set status message
                                }
                                _ => {
                                    // Cell exists but has no history, or cell doesn't exist yet
                                    *status_msg = format!(
                                        "No recorded history for {}",
                                        cell_ref.to_uppercase()
                                    )
                                    .to_string();
                                    // status_msg.push_str(&format!("No recorded history for {}", cell_ref.to_uppercase()));
                                }
                            }
                            sheet.skip_default_display = true; // Don't redisplay grid after history
                        }
                        #[cfg(not(feature = "cell_history"))] // [4]
                        {
                            *status_msg = "Cell history feature is not enabled.".to_string();
                            //  status_msg.push_str("Cell history feature is not enabled.");
                            // sheet.skip_default_display = true; // Prevent grid redraw
                        }
                        // --- End Feature Check ---
                    }
                } else {
                    *status_msg = format!("Invalid cell reference: {}", cell_ref).to_string();
                    // status_msg.push_str(&format!("Invalid cell reference: {}", cell_ref));
                }
            } else {
                *status_msg = "Usage: history <CellReference>".to_string();
            }
        // --- End history command handling ---

        // --- Add undo/redo command handling ---
        } else if cmd == "undo" {
            // --- Feature Check ---
            #[cfg(feature = "undo_state")] // [6, 8, 9]
            {
                sheet.undo(status_msg); // Call the undo method [1]
                                        // status_msg is set within the undo method
            }
            #[cfg(not(feature = "undo_state"))] // [6, 8, 9]
            {
                *status_msg = "Undo feature is not enabled.".to_string();
                //  status_msg.push_str("Undo feature is not enabled.");
            }
            // --- End Feature Check ---
        } else if cmd == "redo" {
            // --- Feature Check ---
            #[cfg(feature = "undo_state")] // <-- Update feature name [1, 3]
            {
                sheet.redo(status_msg); // Call the redo method (sets status_msg) [1]
            }
            #[cfg(not(feature = "undo_state"))] // <-- Update feature name [1, 3]
            {
                *status_msg = "Undo/Redo feature is not enabled.".to_string();
                //   status_msg.push_str("Undo/Redo feature is not enabled.");
            }
            // --- End Feature Check ---
            // --- End undo/redo command handling ---
        } else if cmd.contains('=') {
            if let Some(eq_pos) = cmd.find('=') {
                let cell_name = &cmd[..eq_pos];
                let expr = &cmd[eq_pos + 1..];
                if let Some((row, col)) = cell_name_to_coords(cell_name) {
                    if row < 0 || row >= sheet.total_rows || col < 0 || col >= sheet.total_cols {
                        *status_msg = "Cell out of bounds".to_string();
                    } else {
                        // Call update_cell_formula.
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

    pub fn main() {
        let args: Vec<String> = env::args().collect();
        if args.len() != 3 {
            eprintln!("Usage: {} <rows> <cols>", args[0]);
            return;
        }
        let rows: i32 = args[1].parse().unwrap_or(0);
        let cols: i32 = args[2].parse().unwrap_or(0);
        if rows < 1 || cols < 1 {
            eprintln!("Invalid dimensions.");
            return;
        }
        let mut cmd = String::new();
        let mut status_msg = String::from("ok");
        let mut elapsed_time = 0.0;

        // Allocate the spreadsheet on the heap.
        let mut sheet = Box::new(Spreadsheet::new(rows, cols));
        println!(
            "Boxed sheet at address {:p}, rows={}, cols={}",
            &*sheet, sheet.total_rows, sheet.total_cols
        );

        display_grid(&sheet);
        print!("[{:.1}] ({}) > ", elapsed_time, status_msg);
        io::stdout().flush().unwrap();

        
        
        
        
        
        
        let mut cmd = String::new();
        loop {
            cmd.clear();
            // 1) Read a line, bail out on EOF
            let bytes = match io::stdin().read_line(&mut cmd) {
                Ok(n) => n,
                Err(_) => 0,
            };
            if bytes == 0 {
                // EOF
                break;
            }
        
            let cmd = cmd.trim();
            // explicit quit
            if cmd == "q" {
                break;
            }
        
            // 2) Only treat it as a real command if it matches one of your patterns
            let is_scroll = matches!(cmd, "w" | "a" | "s" | "d");
            let is_jump   = cmd.starts_with("scroll_to ");
            let is_toggle = cmd == "enable_output" || cmd == "disable_output";
            let is_cache  = cmd == "clear_cache";
            let is_history= cmd.contains("history");
            let is_assign = cmd.contains('=');  // crude but works for A1=3, etc.
        
            if !(is_scroll || is_jump || is_toggle || is_cache || is_assign||is_history) {
                // garbage (a stray char), skip it
                continue;
            }
        
            // at this point it’s a real, supported command → process & display
            let start = Instant::now();
            process_command(&mut *sheet, cmd, &mut status_msg);
            elapsed_time = start.elapsed().as_secs_f64();
        
            if sheet.output_enabled {
                display_grid_from(&sheet, sheet.top_row, sheet.left_col);
            }
            print!("[{:.1}] ({}) > ", elapsed_time, status_msg);
            io::stdout().flush().unwrap();
            status_msg = "ok".to_string();
        }
        
    }
}

#[cfg(feature = "gui_app")]
mod gui_app {

    use eframe::egui;
    // Use the modules declared above
    use crate::parser::*; // Correct path
    use crate::sheet::*; // Correct path
                         // --- Add necessary imports ---
    use egui_extras::{Column, Size, StripBuilder, TableBuilder}; // Added Column

    use std::env;
    use std::time::Duration;
    use std::time::Instant;

    // Imports needed for charting and UI
    use egui::ComboBox;
    use egui::Vec2b; // For axis configuration
    use egui_plot::{Bar, BarChart, Legend, Line, Plot, PlotPoints, Points}; // For the dropdown
                                                                            // Add linreg import
    use linreg::linear_regression;
    // Import Color32
    use egui::Color32;

    // --- Define a palette of distinct colors ---
    const PLOT_COLORS: [Color32; 8] = [
        Color32::from_rgb(100, 143, 255), // Blueish
        Color32::from_rgb(250, 120, 120), // Reddish
        Color32::from_rgb(140, 230, 140), // Greenish
        Color32::from_rgb(255, 180, 80),  // Orangey
        Color32::from_rgb(160, 160, 255), // Purplish
        Color32::from_rgb(255, 255, 120), // Yellowish
        Color32::from_rgb(120, 200, 200), // Cyanish
        Color32::from_rgb(220, 140, 220), // Pinkish
                                          // Add more colors if needed
    ];
    // --- Helper functions needed in main.rs scope ---

    /// Converts a 0-indexed column number (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
    /// into its corresponding alphabetical string representation.
    fn col_to_letters(mut col: i32) -> String {
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

    /// Converts (row, col) 0-indexed coordinates to cell name (e.g., "A1").
    /// Based on the implementation in sheet.rs [1].
    fn coords_to_cell_name(row: i32, col: i32) -> String {
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

    // --- Charting Data Structures ---

    // Define an enum for chart types
    #[derive(Debug, PartialEq, Clone)]
    enum ChartType {
        Bar,
        Line,
        Scatter,
    }

    // --- REVISED: Structure for Grouped Bar Chart Data ---
    #[derive(Clone)]
    struct GroupedBarChartData {
        title: String,
        category_names: Vec<String>, // Names for X-axis ticks (from rows)
        // Each tuple is (Series Name, Vec<Value for each category>)
        series: Vec<(String, Vec<f64>)>,
    }

    // --- NEW: Structure to hold scatter plot data ---
    #[derive(Clone)]
    struct ScatterChartData {
        title: String,
        // Store points directly. Could add series name later if multiple series needed.
        points: Vec<[f64; 2]>,
        // Add field to store the two points defining the trendline (start, end)
        trendline_points: Option<Vec<[f64; 2]>>, // Will contain [[x_min, y_at_x_min], [x_max, y_at_x_max]]
                                                 // Optional: Add labels corresponding to points for hover/tooltips later
                                                 // point_labels: Vec<String>,
    }
    // Structure to hold prepared line chart data
    // Stores Vec<[f64; 2]> directly as it's Cloneable
    #[derive(Clone)] // Use derive since Vec<[f64; 2]> is Clone
    struct LineChartData {
        title: String,
        x_labels: Vec<String>,
        lines: Vec<(String, Vec<[f64; 2]>)>, // Store cloneable points data
    }

    // Enum to hold data for different plot types
    #[derive(Clone)] // Use derive since all contained types are Clone
    enum ChartData {
        GroupedBar(GroupedBarChartData),
        Line(LineChartData),
        Scatter(ScatterChartData), // <-- Add Scatter variant
    }

    // --- Application State ---
    struct MyApp {
        spreadsheet: Box<Spreadsheet>, // From sheet.rs [1]
        status_message: String,
        selected_cell: Option<(i32, i32)>,
        formula_input: String,
        last_elapsed_time: f64,

        // Charting State
        show_chart_config_window: bool,
        chart_config_type: ChartType,
        chart_config_title: String,
        chart_error_message: String,

        // // Config for Bar Chart
        // chart_config_range_categories: String,
        // chart_config_range_values: String,

        // Config for Line Chart
        chart_config_range_data: String,
        chart_config_x_labels: Vec<String>,
        chart_config_line_names: Vec<String>,
        chart_config_parsed_dims: Option<(usize, usize)>, // (num_rows, num_cols)

        // --- NEW Config for Scatter Chart ---
        chart_config_range_x_values: String, // e.g., "A1:A10"
        chart_config_range_y_values: String, // e.g., "B1:B10"

        // Chart Display State
        chart_to_display: Option<ChartData>,
        // --- NEW State for Focus ---
        request_focus_formula_bar: bool,
    }

    // --- MyApp Implementation ---
    impl MyApp {
        fn new(cc: &eframe::CreationContext<'_>, rows: i32, cols: i32) -> Self {
            egui::Context::set_visuals(&cc.egui_ctx, egui::Visuals::dark());

            // Spreadsheet::new returns Box<Spreadsheet> [1]
            let mut sheet = Spreadsheet::new(rows, cols);
            sheet.output_enabled = true; // Assuming this field exists in Spreadsheet [1]

            println!(
                "Boxed sheet at address {:p}, rows={}, cols={}",
                &*sheet, sheet.total_rows, sheet.total_cols
            );

            // Fetch initial formula *before* moving sheet into Self
            let initial_formula = sheet.get_cell_raw_content(0, 0); // Call the method below

            Self {
                spreadsheet: sheet,
                status_message: "Ready".to_string(),
                selected_cell: Some((0, 0)),
                formula_input: initial_formula,
                last_elapsed_time: 0.0,

                // Charting State Init
                show_chart_config_window: false,
                chart_config_type: ChartType::Bar,
                chart_config_title: "My Chart".to_string(),
                chart_error_message: String::new(),
                // chart_config_range_categories: "A1:A5".to_string(),
                // chart_config_range_values: "B1:B5".to_string(),
                chart_config_range_data: "A2:C4".to_string(),
                chart_config_x_labels: Vec::new(),
                chart_config_line_names: Vec::new(),
                chart_config_parsed_dims: None,
                chart_to_display: None,
                // --- NEW Scatter Config Init ---
                chart_config_range_x_values: "A1:A10".to_string(), // Example default
                chart_config_range_y_values: "B1:B10".to_string(), // Example default
                request_focus_formula_bar: false,
            }
        }

        // Helper to get raw cell content (implementing the previously discussed function)
        fn get_cell_raw_content(&self, row: i32, col: i32) -> String {
            // Use get_formula from sheet.rs [1]
            self.spreadsheet
                .get_formula(row, col)
                // Fallback to showing the numeric value if no formula exists
                .unwrap_or_else(|| self.spreadsheet.get_cell_value(row, col).to_string())
            // get_cell_value returns i32 [1]
        }

        // Helper to update the formula input when a cell is selected
        fn update_formula_bar_on_select(&mut self) {
            if let Some((r, c)) = self.selected_cell {
                self.formula_input = self.get_cell_raw_content(r, c); // Use the helper above
            } else {
                self.formula_input.clear();
            }
        }

        // Helper to commit the formula from the input bar
        fn commit_formula_input(&mut self) {
            if let Some((row, col)) = self.selected_cell {
                let start = Instant::now();

                // --- Corrected '=' sign handling ---
                let raw_input = self.formula_input.trim();
                let formula_to_evaluate = if raw_input.starts_with('=') {
                    raw_input.get(1..).unwrap_or("").trim_start()
                } else {
                    // Treat as literal if no '=' - parser handles numbers, sheet handles storage
                    raw_input
                };
                // --- End Correction ---

                // Pass the processed formula string
                // Assuming update_cell_formula exists in sheet.rs [1]
                self.spreadsheet.update_cell_formula(
                    row,
                    col,
                    formula_to_evaluate, // <-- Pass the CORRECT variable
                    &mut self.status_message,
                );

                let duration = start.elapsed();
                self.last_elapsed_time = duration.as_secs_f64();

                // Optional: Refresh formula bar *after* commit
                self.update_formula_bar_on_select();
            } else {
                self.status_message = "No cell selected".to_string();
                self.last_elapsed_time = 0.0;
            }
        }

        // Helper: Parse Range string
        fn parse_range(&self, range_str: &str) -> Result<((i32, i32), (i32, i32)), String> {
            let parts: Vec<&str> = range_str.split(':').map(str::trim).collect();
            if parts.len() != 2 {
                return Err(format!("Invalid range format: {}", range_str));
            }
            // cell_name_to_coords is defined in sheet.rs [1]
            let coord1_opt = cell_name_to_coords(parts[0]);
            let coord2_opt = cell_name_to_coords(parts[1]);

            match (coord1_opt, coord2_opt) {
                (Some(coord1_tuple), Some(coord2_tuple)) => {
                    // Check bounds using fields from Spreadsheet [1]
                    if coord1_tuple.0 < 0
                        || coord1_tuple.0 >= self.spreadsheet.total_rows
                        || coord1_tuple.1 < 0
                        || coord1_tuple.1 >= self.spreadsheet.total_cols
                        || coord2_tuple.0 < 0
                        || coord2_tuple.0 >= self.spreadsheet.total_rows
                        || coord2_tuple.1 < 0
                        || coord2_tuple.1 >= self.spreadsheet.total_cols
                    {
                        Err("Range coordinates out of bounds".to_string())
                    } else {
                        // Calculate and return ordered coords
                        let min_row = coord1_tuple.0.min(coord2_tuple.0);
                        let min_col = coord1_tuple.1.min(coord2_tuple.1);
                        let max_row = coord1_tuple.0.max(coord2_tuple.0);
                        let max_col = coord1_tuple.1.max(coord2_tuple.1);
                        Ok(((min_row, min_col), (max_row, max_col)))
                    }
                }
                _ => Err(format!("Invalid cell names in range: {}", range_str)),
            }
        }

        // Helper to update dynamic line chart config fields
        fn update_dynamic_chart_config_dims(&mut self) {
            self.chart_error_message.clear();
            match self.parse_range(&self.chart_config_range_data) {
                Ok(((r1, c1), (r2, c2))) => {
                    let num_rows = (r2 - r1 + 1) as usize;
                    let num_cols = (c2 - c1 + 1) as usize;

                    if num_rows == 0 || num_cols == 0 {
                        self.chart_error_message = "Range cannot be empty".to_string();
                        self.chart_config_parsed_dims = None;
                        return;
                    }
                    self.chart_config_parsed_dims = Some((num_rows, num_cols));

                    // Resize/populate labels (use default row numbers)
                    if self.chart_config_x_labels.len() != num_rows {
                        self.chart_config_x_labels = (0..num_rows)
                            .map(|i| format!("Row {}", r1 + 1 + i as i32))
                            .collect();
                    }
                    // Resize/populate names (use default column letters)
                    if self.chart_config_line_names.len() != num_cols {
                        self.chart_config_line_names = (0..num_cols)
                            .map(|i| col_to_letters(c1 + i as i32))
                            .collect();
                    }
                }
                Err(e) => {
                    self.chart_error_message = e;
                    self.chart_config_parsed_dims = None;
                }
            }
        }

        // Helper: Generate Chart Data
        fn generate_chart_data(&mut self) {
            self.chart_to_display = None; // Clear previous chart
            self.chart_error_message.clear();

            match self.chart_config_type {
                // --- REVISED Bar Chart Logic ---
                ChartType::Bar => {
                    // Ensure dimensions are parsed from the range input
                    if self.chart_config_parsed_dims.is_none() {
                        self.update_dynamic_chart_config_dims(); // Use shared helper
                        if self.chart_config_parsed_dims.is_none() {
                            return;
                        } // Error already set
                    }

                    let range_result = self.parse_range(&self.chart_config_range_data);
                    if let Err(e) = range_result {
                        self.chart_error_message = e;
                        return;
                    }
                    let ((r1, c1), (r2, c2)) = range_result.unwrap();

                    let num_rows = (r2 - r1 + 1) as usize; // Categories
                    let num_cols = (c2 - c1 + 1) as usize; // Series

                    let mut series_data: Vec<(String, Vec<f64>)> = Vec::with_capacity(num_cols);

                    // Fetch Data: Iterate Columns (Series) first
                    for i in 0..num_cols {
                        let current_col = c1 + i as i32;
                        // Get series name from config state
                        let series_name = self
                            .chart_config_line_names
                            .get(i)
                            .cloned()
                            .unwrap_or_else(|| col_to_letters(current_col));

                        let mut values_for_series: Vec<f64> = Vec::with_capacity(num_rows);

                        // Iterate Rows (Categories) for this series
                        for j in 0..num_rows {
                            let current_row = r1 + j as i32;
                            let value =
                                self.spreadsheet.get_cell_value(current_row, current_col) as f64;
                            if self.spreadsheet.get_cell_status(current_row, current_col)
                                == CellStatus::Error
                            {
                                self.chart_error_message = format!(
                                    "Error in cell: {}",
                                    coords_to_cell_name(current_row, current_col)
                                );
                                return;
                            }
                            values_for_series.push(value);
                        }
                        series_data.push((series_name, values_for_series));
                    }

                    // Store as GroupedBar ChartData
                    self.chart_to_display = Some(ChartData::GroupedBar(GroupedBarChartData {
                        title: self.chart_config_title.clone(),
                        // Get category names from config state
                        category_names: self.chart_config_x_labels.clone(),
                        series: series_data,
                    }));
                } // --- End Revised Bar Chart Logic ---
                ChartType::Line => {
                    // Ensure dimensions are parsed
                    if self.chart_config_parsed_dims.is_none() {
                        self.update_dynamic_chart_config_dims();
                        if self.chart_config_parsed_dims.is_none() {
                            return;
                        } // Error msg already set
                    }

                    // Handle Result from parse_range
                    let range_result = self.parse_range(&self.chart_config_range_data);
                    if let Err(e) = range_result {
                        self.chart_error_message = e;
                        return;
                    }
                    let ((r1, c1), (r2, c2)) = range_result.unwrap(); // Safe

                    let num_rows = (r2 - r1 + 1) as usize;
                    let num_cols = (c2 - c1 + 1) as usize;

                    // Store Vec<(String, Vec<[f64; 2]>)> directly
                    let mut lines_data: Vec<(String, Vec<[f64; 2]>)> = Vec::with_capacity(num_cols);

                    // Fetch data (Cols -> Lines, Rows -> Points)
                    for i in 0..num_cols {
                        // Iterate Columns
                        let current_col = c1 + i as i32;
                        let line_name = self
                            .chart_config_line_names
                            .get(i)
                            .cloned()
                            .unwrap_or_else(|| col_to_letters(current_col));

                        let mut points: Vec<[f64; 2]> = Vec::with_capacity(num_rows);

                        for j in 0..num_rows {
                            // Iterate Rows
                            let current_row = r1 + j as i32;
                            let x_value = j as f64; // Use 0-based index for X

                            // get_cell_value returns i32 [1]
                            let y_value = self.spreadsheet.get_cell_value(current_row, current_col);
                            // get_cell_status exists [1]
                            if self.spreadsheet.get_cell_status(current_row, current_col)
                                == CellStatus::Error
                            {
                                self.chart_error_message = format!(
                                    "Error in value cell: {}",
                                    coords_to_cell_name(current_row, current_col)
                                );
                                return;
                            }
                            points.push([x_value, y_value as f64]);
                        }
                        // Store the Vec<[f64; 2]> directly
                        lines_data.push((line_name, points));
                    }

                    // Store result
                    self.chart_to_display = Some(ChartData::Line(LineChartData {
                        title: self.chart_config_title.clone(),
                        x_labels: self.chart_config_x_labels.clone(),
                        lines: lines_data, // Store the cloneable Vec<(String, Vec<[f64; 2]>)>
                    }));
                }
                ChartType::Scatter => {
                    // 1. Parse Ranges (as before)
                    let x_range_result = self.parse_range(&self.chart_config_range_x_values);
                    if let Err(e) = x_range_result {
                        self.chart_error_message = e;
                        return;
                    }
                    let x_range = x_range_result.unwrap();

                    let y_range_result = self.parse_range(&self.chart_config_range_y_values);
                    if let Err(e) = y_range_result {
                        self.chart_error_message = e;
                        return;
                    }
                    let y_range = y_range_result.unwrap();

                    // 2. Validation (as before)
                    let x_len =
                        (x_range.1 .0 - x_range.0 .0 + 1) * (x_range.1 .1 - x_range.0 .1 + 1);
                    if x_len == 0 {
                        /* error */
                        return;
                    }
                    let y_len =
                        (y_range.1 .0 - y_range.0 .0 + 1) * (y_range.1 .1 - y_range.0 .1 + 1);
                    if x_len != y_len {
                        /* error */
                        return;
                    }
                    let x_is_col = x_range.0 .1 == x_range.1 .1;
                    let y_is_col = y_range.0 .1 == y_range.1 .1;

                    // 3. Fetch Data (as before)
                    let mut points: Vec<[f64; 2]> = Vec::with_capacity(x_len as usize);
                    let mut xs: Vec<f64> = Vec::with_capacity(x_len as usize); // For regression
                    let mut ys: Vec<f64> = Vec::with_capacity(x_len as usize); // For regression
                    for i in 0..x_len {
                        let (x_r, x_c) = if x_is_col {
                            (x_range.0 .0 + i, x_range.0 .1)
                        } else {
                            (x_range.0 .0, x_range.0 .1 + i)
                        };
                        let (y_r, y_c) = if y_is_col {
                            (y_range.0 .0 + i, y_range.0 .1)
                        } else {
                            (y_range.0 .0, y_range.0 .1 + i)
                        };

                        let x_value = self.spreadsheet.get_cell_value(x_r, x_c) as f64;
                        if self.spreadsheet.get_cell_status(x_r, x_c) == CellStatus::Error {
                            /* error */
                            return;
                        }
                        let y_value = self.spreadsheet.get_cell_value(y_r, y_c) as f64;
                        if self.spreadsheet.get_cell_status(y_r, y_c) == CellStatus::Error {
                            /* error */
                            return;
                        }

                        points.push([x_value, y_value]);
                        xs.push(x_value);
                        ys.push(y_value);
                    }

                    // --- 4. Calculate Trendline ---
                    let mut trendline_data: Option<Vec<[f64; 2]>> = None;
                    // linear_regression takes slices [6]
                    match linear_regression::<f64, f64, f64>(&xs, &ys) {
                        Ok((slope, intercept)) => {
                            // Find min/max X for the line ends
                            // Use fold for robustness against empty xs (though we check x_len earlier)
                            if let (Some(min_x), Some(max_x)) =
                                xs.iter().fold((None, None), |(min_acc, max_acc), &x| {
                                    let new_min =
                                        min_acc.map_or(Some(x), |min_val| Some(x.min(min_val)));
                                    let new_max =
                                        max_acc.map_or(Some(x), |max_val| Some(x.max(max_val)));
                                    (new_min, new_max)
                                })
                            {
                                // Calculate Y values at the min and max X
                                let y_at_min_x = slope * min_x + intercept;
                                let y_at_max_x = slope * max_x + intercept;
                                // Store the start and end points of the trendline
                                trendline_data =
                                    Some(vec![[min_x, y_at_min_x], [max_x, y_at_max_x]]);
                            } else {
                                self.chart_error_message =
                                    "Could not determine X range for trendline.".to_string();
                            }
                        }
                        Err(err) => {
                            // Regression failed (e.g., insufficient data, vertical line)
                            // Optionally provide more specific message based on linreg::Error type
                            self.chart_error_message =
                                format!("Could not calculate trendline: {:?}", err);
                        }
                    }
                    // --- End Trendline Calculation ---

                    // 5. Store Result
                    self.chart_to_display = Some(ChartData::Scatter(ScatterChartData {
                        title: self.chart_config_title.clone(),
                        points,
                        trendline_points: trendline_data, // Store the calculated trendline
                    }));
                } // --- End Scatter Chart Logic ---
            }
            // Close config window on success
            if self.chart_error_message.is_empty() {
                self.show_chart_config_window = false;
            }
        }
    } // End impl MyApp

    // --- eframe::App Implementation ---
    impl eframe::App for MyApp {
        fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
            // --- Menu Bar ---
            egui::TopBottomPanel::top("menu_panel").show(ctx, |ui| {
                egui::menu::bar(ui, |ui| {
                    ui.menu_button("File", |ui| {
                        if ui.button("Quit").clicked() {
                            ui.ctx().send_viewport_cmd(egui::ViewportCommand::Close);
                        }
                    });
                    ui.menu_button("Insert", |ui| {
                        // --- Rename Bar Button ---
                        if ui.button("Grouped Bar Chart...").clicked() {
                            self.chart_config_type = ChartType::Bar;
                            self.update_dynamic_chart_config_dims(); // Use shared helper
                            self.show_chart_config_window = true;
                            self.chart_to_display = None;
                            // Error potentially set by update_...
                            ui.close_menu();
                        }
                        if ui.button("Line Chart...").clicked() {
                            self.chart_config_type = ChartType::Line;
                            self.update_dynamic_chart_config_dims(); // Try to pre-populate config
                            self.show_chart_config_window = true;
                            self.chart_to_display = None;
                            ui.close_menu();
                        }
                        // --- Add Scatter Button ---
                        if ui.button("Scatter Plot...").clicked() {
                            self.chart_config_type = ChartType::Scatter;
                            // Reset state, show window (no dynamic dims needed for basic scatter)
                            self.show_chart_config_window = true;
                            self.chart_to_display = None;
                            self.chart_error_message.clear();
                            ui.close_menu();
                        }
                    });
                });
            });

            // --- Formula Bar Panel ---
            egui::TopBottomPanel::top("formula_panel").show(ctx, |ui| {
                ui.horizontal(|ui| {
                    let selected_label = match self.selected_cell {
                        Some((r, c)) => coords_to_cell_name(r, c), // Use helper
                        None => "None".to_string(),
                    };
                    ui.label("Selected:");
                    ui.add_sized(
                        [60.0, ui.available_height()],
                        egui::Label::new(&selected_label),
                    );

                    ui.label("fx:");
                    let response = ui.add(
                        egui::TextEdit::singleline(&mut self.formula_input)
                            .desired_width(f32::INFINITY),
                    );
                    // Check the flag AFTER adding the widget
                    if self.request_focus_formula_bar {
                        // Request focus using the widget's response ID [3]
                        ui.memory_mut(|m| m.request_focus(response.id));
                        // Reset the flag immediately after requesting
                        self.request_focus_formula_bar = false;
                    }
                    // --- End focus handling ---

                    if response.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                        self.commit_formula_input();
                    }
                    if ui.button("Set").clicked() {
                        self.commit_formula_input();
                    }
                    if ui.button("Clear Cache").clicked() {
                        let start = Instant::now();
                        // Assume cache field exists [1]
                        self.spreadsheet.cache.clear();
                        // Assume dirty_cells field exists [1]
                        self.spreadsheet.dirty_cells.clear();
                        // Assume clear_range_cache exists in parser.rs [2]
                        clear_range_cache();
                        self.status_message = "Cache cleared".to_string();
                        let duration = start.elapsed();
                        self.last_elapsed_time = duration.as_secs_f64();
                    }
                });
            });

            // --- Status Bar Panel ---
            egui::TopBottomPanel::bottom("bottom_panel").show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.label(format!("Status: {}", self.status_message));
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.label(format!("[{:.1} ms]", self.last_elapsed_time * 1000.0));
                        // Assuming output_enabled field exists [1]
                        // ui.checkbox(&mut self.spreadsheet.output_enabled, "Show Updates"); // Removed as it's internal state now
                    });
                });
            });

            // --- START REPLACEMENT: Central Panel with TableBuilder ---
            egui::CentralPanel::default().show(ctx, |ui| {
                // Estimate row height - adjust as needed, e.g., based on font size
                let text_height = egui::TextStyle::Body.resolve(ui.style()).size;
                // Add padding (e.g., for cell borders/spacing)
                // Consider making this slightly larger than the minimal text height
                let row_height = text_height + 4.0; // Example padding

                // --- Use TableBuilder for efficient virtualized grid ---
                TableBuilder::new(ui)
                    .striped(true) // Alternating row colors
                    .resizable(true) // Allow column resizing by dragging
                    // --- FIX 1: Wrap Size in Column::new() ---
                    // Define Row Header column using Column::exact()
                    .column(Column::exact(40.0))
                    // Define Data Columns using Column::initial().at_least()
                    .columns(
                        // The template Column for data columns
                        Column::initial(80.0).at_least(30.0),
                        self.spreadsheet.total_cols as usize, // Number of data columns
                    )
                    // --- End FIX 1 ---
                    .header(20.0, |mut header| {
                        // Header row height
                        // --- Column Headers ---
                        header.col(|ui| {
                            ui.strong("");
                        }); // Top-left corner empty
                        for c in 0..self.spreadsheet.total_cols {
                            header.col(|ui| {
                                // Display column letters (A, B, C...)
                                ui.strong(col_to_letters(c));
                            });
                        }
                    })
                    .body(|mut body| {
                        // --- FIX 2: Correct closure signature and get index ---
                        body.rows(
                            row_height,
                            self.spreadsheet.total_rows as usize,
                            // Closure now takes only `mut row`
                            |mut row| {
                                // Get the row index from the TableRow object [5]
                                let row_index = row.index();
                                let r = row_index as i32; // Convert to i32 for your logic

                                // Row Header (No change needed inside)
                                row.col(|ui| {
                                    ui.label(format!("{}", r + 1));
                                });

                                // Cells (No change needed inside loop)
                                for c in 0..self.spreadsheet.total_cols {
                                    row.col(|ui| {
                                        let is_selected = self.selected_cell == Some((r, c));
                                        let cell_status = self.spreadsheet.get_cell_status(r, c);
                                        let cell_value_str = if cell_status == CellStatus::Error {
                                            "ERR".to_string()
                                        } else {
                                            self.spreadsheet.get_cell_value(r, c).to_string()
                                        };
                                        let response = ui.add_sized(
                                            ui.available_size(),
                                            egui::SelectableLabel::new(is_selected, cell_value_str),
                                        );
                                        if response.clicked() {
                                            let new_selection = Some((r, c));
                                            if self.selected_cell != new_selection {
                                                self.selected_cell = new_selection;
                                                self.update_formula_bar_on_select();
                                                self.request_focus_formula_bar = true;
                                                self.status_message = "ok".to_string();
                                                self.last_elapsed_time = 0.0;
                                            }
                                        }
                                    }); // End cell column closure
                                } // End column loop
                            }, // End row closure
                        ); // End body.rows
                           // --- End FIX 2 ---
                    }); // End body
            }); // End CentralPanel
                // --- END REPLACEMENT ---

            // --- Chart Configuration Window ---
            if self.show_chart_config_window {
                let mut is_open = true;
                egui::Window::new("Chart Configuration")
                    .open(&mut is_open)
                    .resizable(true)
                    .scroll2([false, true]) // Vertical scroll
                    .default_width(350.0)
                    .show(ctx, |ui| {
                        ui.label("Chart Title:");
                        ui.text_edit_singleline(&mut self.chart_config_title);
                        ui.separator();

                        // Chart Type Selection
                        let chart_type_changed = ui
                            .horizontal(|ui| {
                                let mut changed = false;
                                ui.label("Chart Type:");
                                // --- Update Bar Text ---
                                changed |= ui
                                    .selectable_value(
                                        &mut self.chart_config_type,
                                        ChartType::Bar,
                                        "Grouped Bar",
                                    )
                                    .changed();
                                changed |= ui
                                    .selectable_value(
                                        &mut self.chart_config_type,
                                        ChartType::Line,
                                        "Line",
                                    )
                                    .changed();
                                // --- Add Scatter Option ---
                                changed |= ui
                                    .selectable_value(
                                        &mut self.chart_config_type,
                                        ChartType::Scatter,
                                        "Scatter",
                                    )
                                    .changed();
                                changed
                            })
                            .inner;
                        if chart_type_changed
                            && (self.chart_config_type == ChartType::Bar
                                || self.chart_config_type == ChartType::Line)
                        {
                            self.update_dynamic_chart_config_dims();
                        }
                        ui.separator();

                        // Type-Specific Config
                        match self.chart_config_type {
                            // --- Revised Bar Config UI ---
                            ChartType::Bar => {
                                ui.label("Data Range (e.g., A2:C4):");
                                if ui
                                    .text_edit_singleline(&mut self.chart_config_range_data)
                                    .changed()
                                {
                                    self.update_dynamic_chart_config_dims(); // Use shared helper
                                }
                                // Show dynamic fields for category/series names (like Line)
                                if let Some((num_rows, num_cols)) = self.chart_config_parsed_dims {
                                    ui.separator();
                                    ui.label("Category Names (Rows):"); // Renamed from X-Axis
                                    if self.chart_config_x_labels.len() == num_rows {
                                        egui::ScrollArea::vertical()
                                            .id_source("bar_cat_scroll")
                                            .max_height(100.0)
                                            .show(ui, |ui| {
                                                for i in 0..num_rows {
                                                    ui.horizontal(|ui| {
                                                        ui.label(format!("Category {}:", i + 1));
                                                        ui.text_edit_singleline(
                                                            &mut self.chart_config_x_labels[i],
                                                        );
                                                    });
                                                }
                                            });
                                    }
                                    ui.separator();
                                    ui.label("Series Names (Columns):"); // Renamed from Line Names
                                    if self.chart_config_line_names.len() == num_cols {
                                        egui::ScrollArea::vertical()
                                            .id_source("bar_ser_scroll")
                                            .max_height(100.0)
                                            .show(ui, |ui| {
                                                for i in 0..num_cols {
                                                    ui.horizontal(|ui| {
                                                        ui.label(format!("Col {}:", i + 1)); // Label based on index
                                                        ui.text_edit_singleline(
                                                            &mut self.chart_config_line_names[i],
                                                        );
                                                    });
                                                }
                                            });
                                    }
                                } else {
                                    ui.label("(Enter a valid data range above)");
                                }
                            } // --- End Revised Bar Config UI ---
                            ChartType::Line => {
                                ui.label("Data Range (e.g., A2:C4):");
                                if ui
                                    .text_edit_singleline(&mut self.chart_config_range_data)
                                    .changed()
                                {
                                    self.update_dynamic_chart_config_dims();
                                }
                                if let Some((num_rows, num_cols)) = self.chart_config_parsed_dims {
                                    ui.separator();
                                    ui.label("X-Axis Point Names (Rows):");
                                    if self.chart_config_x_labels.len() == num_rows {
                                        egui::ScrollArea::vertical()
                                            .id_source("x_axis_label_scroll")
                                            .max_height(100.0)
                                            .show(ui, |ui| {
                                                for i in 0..num_rows {
                                                    ui.horizontal(|ui| {
                                                        ui.label(format!("Row {}:", i + 1)); // Adjust label if needed
                                                        ui.text_edit_singleline(
                                                            &mut self.chart_config_x_labels[i],
                                                        );
                                                    });
                                                }
                                            });
                                    }
                                    ui.separator();
                                    ui.label("Line/Series Names (Columns):");
                                    if self.chart_config_line_names.len() == num_cols {
                                        egui::ScrollArea::vertical()
                                            .id_source("line_name_scroll")
                                            .max_height(100.0)
                                            .show(ui, |ui| {
                                                for i in 0..num_cols {
                                                    ui.horizontal(|ui| {
                                                        ui.label(format!("Col {}:", i + 1)); // Adjust label if needed
                                                        ui.text_edit_singleline(
                                                            &mut self.chart_config_line_names[i],
                                                        );
                                                    });
                                                }
                                            });
                                    }
                                } else {
                                    ui.label("(Enter a valid data range above)");
                                }
                            }
                            // --- Add Scatter Config UI ---
                            ChartType::Scatter => {
                                ui.label("X-Values Range (e.g., A1:A10):");
                                ui.text_edit_singleline(&mut self.chart_config_range_x_values);
                                ui.label("Y-Values Range (e.g., B1:B10):");
                                ui.text_edit_singleline(&mut self.chart_config_range_y_values);
                                // Optional: Add input for point labels range later
                            }
                        }
                        ui.separator();
                        if !self.chart_error_message.is_empty() {
                            ui.colored_label(egui::Color32::RED, &self.chart_error_message);
                        }
                        ui.add_space(10.0);
                        if ui.button("Generate Chart").clicked() {
                            self.generate_chart_data();
                        }
                    }); // End Window
                if !is_open {
                    self.show_chart_config_window = false;
                    self.chart_error_message.clear();
                }
            }

            // --- Chart Display Window ---
            let mut close_chart_display = false;
            if let Some(chart_data) = &self.chart_to_display {
                let chart_data_clone = chart_data.clone(); // Clone for use in closures
                let mut is_display_open = true;

                egui::Window::new(match &chart_data_clone {
                   ChartData::GroupedBar(data) => &data.title, // Use GroupedBar title
                   ChartData::Line(line_data) => &line_data.title,
                   ChartData::Scatter(scatter_data) => &scatter_data.title, // <-- Add Scatter title
               })
                .open(&mut is_display_open)
                .resizable(true)
                .default_width(500.0)
                .default_height(350.0)
                .show(ctx, |ui| {

                    // --- Create the Plot (mutably) ---
                    let mut plot = Plot::new("chart_plot") // <-- Make `plot` mutable
                        .legend(Legend::default())
                        .auto_bounds_y();

                    // --- Conditionally Apply x_axis_formatter ---
                    match &chart_data_clone {
                        ChartData::Line(line_data) => {
                            let x_labels = line_data.x_labels.clone(); // Clone labels for closure
                            // Apply the formatter to the mutable plot instance
                            plot = plot.x_axis_formatter(move |grid_mark, _max_chars, _range| {
                                let index = grid_mark.value.round() as usize;
                                // Safely get label, fallback to number if index is out of bounds
                                x_labels.get(index).cloned().unwrap_or_else(|| format!("{:.0}", grid_mark.value))
                            });
                        }
                        | // --- Ensure Formatter for GroupedBar ---
                        ChartData::GroupedBar(data) => {
                            let cat_names = data.category_names.clone();
                            plot = plot.x_axis_formatter(move |grid_mark, _, _| {
                                let index = grid_mark.value.round() as usize;
                                cat_names.get(index).cloned().unwrap_or_default()
                            });
                        }
                        // --- Ensure Formatter for Line ---
                        | ChartData::Scatter { .. } => {
                            // No specific formatter needed for Bar chart in this case
                            // Plot remains as initially configured
                            plot = plot.auto_bounds_x();
                        }
                    }
                    // --- End Conditional Modification ---


                    // --- Show the plot and add elements ---
                    // `plot` now has the formatter applied (or not) based on the match above
                    plot.show(ui, |plot_ui| {
                        match &chart_data_clone {
                            // --- Add GroupedBar Plotting ---
                            ChartData::GroupedBar(data) => {
                                let num_series = data.series.len();
                                let num_categories = data.category_names.len();
                                if num_categories == 0 || num_series == 0 { return; } // Nothing to plot

                                // Calculate width for each bar within a group
                                // Make total width slightly less than 1.0 for spacing between groups
                                let total_group_width = 0.8;
                                let bar_width = total_group_width / num_series as f64;

                                // Loop through each SERIES (column)
                                for (series_idx, (series_name, values)) in data.series.iter().enumerate() {
                                    let mut series_bars: Vec<Bar> = Vec::with_capacity(num_categories);
                                    // --- Get color from the palette using modulo ---
                                    let color = PLOT_COLORS[series_idx % PLOT_COLORS.len()];
                                    // --- End color selection ---

                                    // Loop through each CATEGORY (row) for this series
                                    for (cat_idx, value) in values.iter().enumerate() {
                                        // Calculate the center X position for this specific bar within the group
                                        // `cat_idx` is the center of the group (0, 1, 2...)
                                        // Offset based on series index and bar width
                                        let center_offset = (series_idx as f64 - (num_series as f64 - 1.0) / 2.0) * bar_width;
                                        let x_pos = cat_idx as f64 + center_offset;

                                        series_bars.push(
                                            Bar::new(x_pos, *value)
                                                .width(bar_width)
                                                .name(format!("{}: {}", series_name, value)) // Hover text
                                                // Individual color is set on the BarChart below
                                        );
                                    }
                                    // Create a BarChart for THIS series with its color
                                    let bar_chart = BarChart::new(series_bars)
                                                        .name(series_name) // Legend name
                                                        .color(color);
                                    plot_ui.bar_chart(bar_chart);
                                }
                            } // --- End GroupedBar Plotting ---
                            ChartData::Line(line_data) => {
                                for (name, points_vec) in &line_data.lines {
                                    let owned_points_vec = points_vec.clone();
                                    let plot_points = PlotPoints::from(owned_points_vec);
                                    let line = Line::new(plot_points).name(name);
                                    plot_ui.line(line);
                                }
                            }
                            ChartData::Scatter(scatter_data) => {
                                // --- Plot Scatter Points ---
                                let plot_points = PlotPoints::from(scatter_data.points.clone());
                                let points_item = Points::new(plot_points)
                                    .radius(3.0)
                                    .name(&scatter_data.title); // Use main title or specific series name
                                plot_ui.points(points_item);

                                // --- Plot Trendline (If Available) ---
                                if let Some(trend_points_vec) = &scatter_data.trendline_points {
                                    // Convert trendline points (Vec<[f64; 2]>) to PlotPoints
                                    let trend_plot_points = PlotPoints::from(trend_points_vec.clone());
                                    // Create Line item for trendline
                                    let trend_line = Line::new(trend_plot_points)
                                        .color(egui::Color32::RED) // Make trendline distinct
                                        // .style(egui_plot::LineStyle::dashed_dense()) // Optional: dashed style
                                        .name("Trendline"); // Name for legend
                                    // Add line to plot
                                    plot_ui.line(trend_line);
                                }
                                // --- End Trendline Plotting ---
                            }
                        }
                    }); // End plot.show
                }); // End Window

                if !is_display_open {
                    close_chart_display = true;
                }
            }
            if close_chart_display {
                self.chart_to_display = None;
            }

            // Request repaint periodically
            ctx.request_repaint_after(Duration::from_millis(100));
        } // End fn update
    } // End impl eframe::App

    // --- Main function ---
    pub fn main() -> Result<(), eframe::Error> {
        let mut args: Vec<String> = env::args().collect();
        if args.len() != 3 {
            eprintln!("Usage: {} <rows> <cols>", args[0]);
            eprintln!("Using default: 100 rows, 26 cols"); // Adjusted default
            args = vec![args[0].clone(), "100".to_string(), "26".to_string()];
        }
        let mut rows: i32 = args[1].parse().unwrap_or(100);
        let mut cols: i32 = args[2].parse().unwrap_or(26);

        if rows < 1 || cols < 1 {
            eprintln!("Invalid dimensions. Using defaults (100x26).");
            rows = 100;
            cols = 26;
        }

        let options = eframe::NativeOptions {
            // Use the viewport builder to set window properties [9]
            viewport: egui::ViewportBuilder::default()
                // Set the desired initial inner size of the window's client area
                .with_inner_size([1024.0, 768.0]), // You can also set other properties here, e.g.:
            // .with_min_inner_size([400.0, 300.0])
            // .with_title("My Spreadsheet App") // Title can also be set here
            // <-- Comma needed here
            ..Default::default()
        };

        eframe::run_native(
            "Rust Spreadsheet GUI", // Window title
            options,
            Box::new(move |cc| Box::new(MyApp::new(cc, rows, cols))),
        )
    }

    // --- Ensure CellStatus enum is accessible (can be defined here or use sheet::CellStatus) ---
    // Based on sheet.rs [1], it should be accessed via sheet::CellStatus
    // No need to redefine here if mod sheet; is correct.
}

// --- Top-Level Main Dispatcher ---
fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Ensure only one feature is enabled (optional but recommended)
    #[cfg(all(feature = "cli_app", feature = "gui_app"))]
    compile_error!(
        "Features 'cli_app' and 'gui_app' are mutually exclusive. Please enable only one."
    );

    // Dispatch to the correct main function
    #[cfg(feature = "cli_app")]
    {
        cli_app::main();
        Ok(()) // cli_app::main doesn't return Result, so wrap it
    }

    #[cfg(feature = "gui_app")]
    {
        // gui_app::main returns Result<(), eframe::Error>
        // Map the error type to Box<dyn Error> for consistency
        gui_app::main().map_err(|e| Box::new(e) as Box<dyn std::error::Error>)
    }

    // Handle case where neither feature is enabled
    #[cfg(not(any(feature = "cli_app", feature = "gui_app")))]
    {
        eprintln!("Error: No application feature ('cli_app' or 'gui_app') enabled.");
        eprintln!("Build with --features cli_app or --features gui_app");
        // Return an error exit code
        std::process::exit(1);
        // Alternatively, use compile_error! to fail the build:
        // compile_error!("Please enable either the 'cli_app' or 'gui_app' feature.");
    }
}
