//! A tiny, embeddable spreadsheet engine in pure Rust.
//! 
//! # Features
//! 
//! - Sparse storage for large grids
//! - Formula evaluation with `SUM`, `IF`, `COUNTIF`, etc.
//! - Multi-level undo/redo and cell history (behind features)
//! # Rust Spreadsheet Lab
//!
//! # spreadsheet
//!
//! A tiny, embeddable spreadsheet engine in pure Rust.
//! 
//! ## Features
//! - Sparse storage
//! - Formula evaluation (`SUM`, `IF`, `COUNTIF`, …)
//! - Optional undo/redo & cell history
//! This crate implements the COP290 spreadsheet in safe, idiomatic Rust,
//! with both CLI and optional GUI front-ends, formula parsing, undo/redo,
//! history logging, and an interactive chart GUI.
//!
//! ## Core Modules
//!  - [`parser`] — parse and evaluate formulas  
//!  - [`sheet`]  — manage cells, dependencies, and recalculation  
//!  - [`cli_app`] / [`gui_app`] — front-ends for terminal and GUI  
#![allow(warnings)]
pub mod parser;
/// The `parser` module handles lexical analysis, recursive-descent parsing,
/// and evaluation of spreadsheet formulas (SUM, IF, COUNTIF, etc.).
/// Public API:
/// - `evaluate_formula`  
/// - `clear_range_cache`  
/// - `invalidate_cache_for_cell`
pub mod sheet;
/// The `sheet` module manages the grid of [`Cell`]s, dependency graphs,
/// incremental recalculation (topological sort), undo/redo stacks,
/// and viewport scrolling.
// Export the CLI functions for tests to use
#[cfg(feature = "cli_app")]
pub mod cli_app {
     //! A lightweight CLI front-end for testing and autograder compatibility.
    //! Exposes viewport scrolling, cell assignment, cache clearing,
    //! undo/redo, and history commands.
    use crate::parser::*;
    use crate::sheet::*;
    /// ```rust
/// // import from your crate
/// use spreadsheet::cli_app::col_to_letters;
///
/// assert_eq!(col_to_letters(0), "A");
/// assert_eq!(col_to_letters(26), "AA");
    pub fn col_to_letters(mut col: i32) -> String {
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
    /// Clamp the vertical viewport start row to [0, total_rows − height]
    ///
    /// On out-of-bounds, pulls the view back by 10 or to zero.
    pub fn clamp_viewport_ve(total_rows: i32, start_row: &mut i32) {
        if *start_row > total_rows {
            *start_row -= 10;
        } else if *start_row > (total_rows - 10) {
            *start_row = total_rows - 10;
        } else if *start_row < 0 {
            *start_row = 0;
        }
    }
    /// Clamp a horizontal viewport coordinate so it stays within `[0..max_col]`.
    /// Clamp a column index so it never runs off the left or right of the sheet.
    ///
    /// # Parameters
    /// - `total_cols`: the total number of columns in the sheet.
    /// - `start_col`: the mutable column index to clamp in place.
    ///
    /// # Examples
    ///i
    /// ```rust
    /// # use spreadsheet::cli_app::clamp_viewport_hz;
    /// let mut c = 95;
    /// clamp_viewport_hz(90, &mut c);
    /// assert_eq!(c, 85);
    /// ```
    pub fn clamp_viewport_hz(total_cols: i32, start_col: &mut i32) {
        if *start_col > total_cols {
            *start_col -= 10;
        } else if *start_col > (total_cols - 10) {
            *start_col = total_cols - 10;
        } else if *start_col < 0 {
            *start_col = 0;
        }
    }
    /// Process a single user command string, updating `sheet` and `status_msg`.
    ///
    /// Recognized commands:
    /// - `w`, `a`, `s`, `d`: scroll viewport  
    /// - `scroll_to <CELL>`: jump viewport  
    /// - `disable_output` / `enable_output`  
    /// - `clear_cache`  
    /// - `undo` / `redo` (feature-gated)  
    /// - `<CELL>=<EXPR>`: assign formula to a cell  
    /// - `history <CELL>` (feature-gated)
    pub fn process_command(sheet: &mut Box<Spreadsheet>, cmd: &str, status_msg: &mut String) {
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
        } else if cmd.starts_with("history") {
            let parts: Vec<&str> = cmd.split_whitespace().collect();
            if parts.len() == 2 {
                #[cfg(feature = "cell_history")]
                {
                    sheet.skip_default_display = true;
                    *status_msg = "History displayed".to_string();
                }
                #[cfg(not(feature = "cell_history"))]
                {
                    *status_msg = "Cell history feature is not enabled.".to_string();
                }
            }
        } else {
            *status_msg = "unrecognized cmd".to_string();
        }
    }
}
/// The `gui_app` module implements a GUI front-end using `egui`
/// and `eframe`. It provides a spreadsheet grid, formula bar,
/// and charting capabilities. The GUI is built on top of the
/// [`sheet`] and [`parser`] modules, allowing for interactive
/// editing and visualization of spreadsheet data.
#[cfg(feature = "gui_app")]
pub mod gui_app {
    use crate::parser::*;
    use crate::sheet::*;
    /// Convert a 0-indexed column number into its Excel-style letters
    // Implement GUI-related functions needed for testing
    pub fn col_to_letters(mut col: i32) -> String {
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
    /// Turn (row, col) into `"A1"`, `"B3"`, etc.
    ///
    /// # Examples
    ///
    /// ```
    /// assert_eq!(gui_app::coords_to_cell_name(0, 0), "A1");
    /// ```
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
}

// ─── at the very bottom of src/lib.rs ─────────────────────────────────────────
#[cfg(test)]
mod lib_tests {
    // bring in both CLI and GUI
    use super::*;
    use crate::sheet::Spreadsheet;

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_col_to_letters_cli() {
        assert_eq!(cli_app::col_to_letters(0), "A");
        assert_eq!(cli_app::col_to_letters(25), "Z");
        assert_eq!(cli_app::col_to_letters(26), "AA");
        assert_eq!(cli_app::col_to_letters(701), "ZZ");
        assert_eq!(cli_app::col_to_letters(702), "AAA");
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_clamp_viewport_cli() {
        // vertical: total_rows = 40, viewport height = 10 → only subtracts 10 once
        let mut top = 50;
        cli_app::clamp_viewport_ve(40, &mut top);
        assert_eq!(top, 40);

        let mut too_low = -5;
        cli_app::clamp_viewport_ve(100, &mut too_low);
        assert_eq!(too_low, 0);

        // horizontal: total_cols = 90, viewport width = 10 → only subtracts 10 once
        let mut left = 95;
        cli_app::clamp_viewport_hz(90, &mut left);
        assert_eq!(left, 85);

        let mut too_left = -1;
        cli_app::clamp_viewport_hz(10, &mut too_left);
        assert_eq!(too_left, 0);
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_process_command_wasd() {
        let mut sheet = Box::new(Spreadsheet::new(100, 100));
        sheet.top_row = 20;
        sheet.left_col = 30;
        let mut msg = String::new();

        cli_app::process_command(&mut sheet, "w", &mut msg);
        assert_eq!(sheet.top_row, 10);
        cli_app::process_command(&mut sheet, "s", &mut msg);
        assert_eq!(sheet.top_row, 20);
        cli_app::process_command(&mut sheet, "a", &mut msg);
        assert_eq!(sheet.left_col, 20);
        cli_app::process_command(&mut sheet, "d", &mut msg);
        assert_eq!(sheet.left_col, 30);
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_process_scroll_to() {
        let mut sheet = Box::new(Spreadsheet::new(5, 5));
        let mut msg = String::new();

        // valid
        cli_app::process_command(&mut sheet, "scroll_to A3", &mut msg);
        assert_eq!((sheet.top_row, sheet.left_col), (2, 0));
        assert!(msg.is_empty());

        // out of bounds row
        cli_app::process_command(&mut sheet, "scroll_to Z9", &mut msg);
        assert!(msg.contains("out of bounds"));

        // invalid token
        msg.clear();
        cli_app::process_command(&mut sheet, "scroll_to foo", &mut msg);
        assert!(msg.contains("Invalid cell"));

        // malformed
        msg.clear();
        cli_app::process_command(&mut sheet, "scroll_to", &mut msg);
        assert!(msg.contains("Invalid command"));
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_enable_disable_clear_cache() {
        let mut sheet = Box::new(Spreadsheet::new(2, 2));
        let mut msg = String::new();

        // disable/enable output
        sheet.output_enabled = true;
        cli_app::process_command(&mut sheet, "disable_output", &mut msg);
        assert!(!sheet.output_enabled);
        cli_app::process_command(&mut sheet, "enable_output", &mut msg);
        assert!(sheet.output_enabled);

        // clear_cache
        sheet.cache.insert(
            "X".into(),
            crate::sheet::CachedRange {
                value: 1,
                dependencies: std::collections::HashSet::new(),
            },
        );
        sheet.dirty_cells.insert((0, 0));
        cli_app::process_command(&mut sheet, "clear_cache", &mut msg);
        assert_eq!(msg, "Cache cleared");
        assert!(sheet.cache.is_empty());
        assert!(sheet.dirty_cells.is_empty());
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_undo_redo_placeholders() {
        let mut sheet = Box::new(Spreadsheet::new(1, 1));
        let mut msg = String::new();

        // undo/redo without feature
        cli_app::process_command(&mut sheet, "undo", &mut msg);
        assert!(msg.contains("not enabled"));
        cli_app::process_command(&mut sheet, "redo", &mut msg);
        assert!(msg.contains("not enabled"));
    }

    #[test]
    #[cfg(feature = "cli_app")]
    fn test_assignment_and_history() {
        let mut sheet = Box::new(Spreadsheet::new(3, 3));
        let mut msg = String::new();

        // assignment must not panic; we don't care about msg here
        cli_app::process_command(&mut sheet, "B2=123", &mut msg);

        // now check history (feature off)
        msg.clear();
        cli_app::process_command(&mut sheet, "history A1", &mut msg);
        assert!(msg.contains("not enabled"));
    }

    // now GUI side
    #[test]
    #[cfg(feature = "gui_app")]
    fn test_col_to_letters_gui() {
        assert_eq!(gui_app::col_to_letters(0), "A");
        assert_eq!(gui_app::col_to_letters(25), "Z");
        assert_eq!(gui_app::col_to_letters(26), "AA");
    }

    #[test]
    #[cfg(feature = "gui_app")]
    fn test_coords_to_cell_name_gui() {
        assert_eq!(gui_app::coords_to_cell_name(0, 0), "A1");
        assert_eq!(gui_app::coords_to_cell_name(4, 25), "Z5");
        assert_eq!(gui_app::coords_to_cell_name(9, 26), "AA10");
    }
}
