use crate::sheet::{CellStatus, Spreadsheet, CloneableSheet, CachedRange};
use std::thread::sleep;
use std::time::Duration;
use std::collections::{HashMap, HashSet};

// Define the AST node enum for formula parsing
#[derive(Clone, Debug)]
enum ASTNode {
    Literal(i32),
    CellRef(i32, i32),
    BinaryOp(char, Box<ASTNode>, Box<ASTNode>),
    RangeFunction(String, String), // Function name and range string
    SleepFunction(Box<ASTNode>),
}

// Keep the cache in thread_local storage for thread safety
thread_local! {
    static RANGE_CACHE: std::cell::RefCell<HashMap<String, (i32, HashSet<(i32, i32)>)>> = 
        std::cell::RefCell::new(HashMap::new());
}

fn skip_spaces(input: &mut &str) {
    while let Some(ch) = input.chars().next() {
        if ch.is_whitespace() {
            *input = &input[ch.len_utf8()..];
        } else {
            break;
        }
    }
}

fn evaluate_range_function<'a>(sheet: &CloneableSheet<'a>, func_name: &str, range_str: &str, error: &mut i32) -> i32 {
    // Check if we have this range cached
    let cache_key = format!("{}({})", func_name, range_str);
    
    // Try to get from thread-local cache with improved validation
    if let Some((cached_value, _)) = RANGE_CACHE.with(|cache| {
        cache.borrow().get(&cache_key).map(|(val, deps)| (*val, deps.clone()))
    }) {
        return cached_value;
    }

    if let Some(colon_pos) = range_str.find(':') {
        let cell1 = range_str[..colon_pos].trim();
        let cell2 = range_str[colon_pos + 1..].trim();
        let (start_row, start_col) = match crate::sheet::cell_name_to_coords(cell1) {
            Some(coords) => coords,
            None => { *error = 1; return 0; }
        };
        let (end_row, end_col) = match crate::sheet::cell_name_to_coords(cell2) {
            Some(coords) => coords,
            None => { *error = 1; return 0; }
        };
        if start_row > end_row || start_col > end_col {
            *error = 2;
            return 0;
        }
        
        // Check bounds
        if start_row < 0 || end_row >= sheet.total_rows() || 
           start_col < 0 || end_col >= sheet.total_cols() {
            *error = 4;
            return 0;
        }
        
        // For very large ranges, use streaming calculation
        let cell_count = (end_row - start_row + 1) * (end_col - start_col + 1);
        // let use_streaming = cell_count > 1000000;
        
        // // Optimized aggregation for large ranges
        // if use_streaming {
        //     return evaluate_large_range(sheet, func_name, start_row, start_col, end_row, end_col, error, &cache_key);
        // }
        
        // Standard calculation for small to medium ranges
        let mut sum: i64 = 0;
        let mut min_val = i32::MAX;
        let mut max_val = i32::MIN;
        let mut count = 0;
        let mut dependencies = HashSet::new();

        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if let Some(cell) = sheet.get_cell(r, c) {
                    if cell.status == CellStatus::Error {
                        *error = 3;
                        return 0;
                    }
                    dependencies.insert((r, c));
                    let value = cell.value;
                    sum += value as i64;
                    if value < min_val { min_val = value; }
                    if value > max_val { max_val = value; }
                    count += 1;
                }
            }
        }
        
        if count == 0 {
            *error = 1;
            return 0;
        }
        
        let result = match func_name {
            "MIN" => min_val,
            "MAX" => max_val,
            "SUM" => sum as i32,
            "AVG" => (sum / (count as i64)) as i32,
            "STDEV" => {
                let mean = (sum as f64) / (count as f64);
                let mut variance = 0.0;
                for r in start_row..=end_row {
                    for c in start_col..=end_col {
                        if let Some(cell) = sheet.get_cell(r, c) {
                            let diff = (cell.value as f64) - mean;
                            variance += diff * diff;
                        }
                    }
                }
                variance /= count as f64;
                (variance.sqrt()).round() as i32
            },
            _ => { *error = 1; 0 }
        };
         // Cache the result with full dependencies for smaller ranges
         RANGE_CACHE.with(|cache| {
            cache.borrow_mut().insert(cache_key, (result, dependencies));
        });
        
        // // For large ranges, store minimal dependency information
        // if dependencies.len() > 1000000 {
        //     // Just store the corners and the result to save memory
        //     let mut minimal_deps = HashSet::new();
        //     minimal_deps.insert((start_row, start_col));
        //     minimal_deps.insert((start_row, end_col));
        //     minimal_deps.insert((end_row, start_col));
        //     minimal_deps.insert((end_row, end_col));
            
        //     // Cache the result with minimal dependencies
        //     RANGE_CACHE.with(|cache| {
        //         cache.borrow_mut().insert(cache_key, (result, minimal_deps));
        //     });
        // } else {
        //     // Cache the result with full dependencies for smaller ranges
        //     RANGE_CACHE.with(|cache| {
        //         cache.borrow_mut().insert(cache_key, (result, dependencies));
        //     });
        // }
        
        result
    } else {
        *error = 1;
        0
    }
}

// New function to handle large ranges more efficiently
fn evaluate_large_range<'a>(
    sheet: &CloneableSheet<'a>, 
    func_name: &str,
    start_row: i32, 
    start_col: i32, 
    end_row: i32, 
    end_col: i32,
    error: &mut i32,
    cache_key: &str
) -> i32 {
    // Process in chunks to avoid excessive memory usage
    const CHUNK_SIZE: i32 = 128;
    
    let mut sum: i64 = 0;
    let mut min_val = i32::MAX;
    let mut max_val = i32::MIN;
    let mut count = 0;
    let mut sum_squares: f64 = 0.0;
    
    // For very large ranges, we'll compute statistics in a single pass
    for chunk_row in (start_row..=end_row).step_by(CHUNK_SIZE as usize) {
        let chunk_end_row = (chunk_row + CHUNK_SIZE - 1).min(end_row);
        
        for chunk_col in (start_col..=end_col).step_by(CHUNK_SIZE as usize) {
            let chunk_end_col = (chunk_col + CHUNK_SIZE - 1).min(end_col);
            
            // Process this chunk
            for r in chunk_row..=chunk_end_row {
                for c in chunk_col..=chunk_end_col {
                    if let Some(cell) = sheet.get_cell(r, c) {
                        if cell.status == CellStatus::Error {
                            *error = 3;
                            return 0;
                        }
                        
                        let value = cell.value;
                        sum += value as i64;
                        sum_squares += (value as f64) * (value as f64);
                        
                        if value < min_val { min_val = value; }
                        if value > max_val { max_val = value; }
                        count += 1;
                    }
                }
            }
        }
    }
    
    if count == 0 {
        *error = 1;
        return 0;
    }
    
    // Calculate the result based on function
    let result = match func_name {
        "MIN" => min_val,
        "MAX" => max_val,
        "SUM" => {
            if sum > i32::MAX as i64 || sum < i32::MIN as i64 {
                *error = 3;  // Overflow
                return 0;
            }
            sum as i32
        },
        "AVG" => {
            let avg = sum / (count as i64);
            if avg > i32::MAX as i64 || avg < i32::MIN as i64 {
                *error = 3;  // Overflow
                return 0;
            }
            avg as i32
        },
        "STDEV" => {
            let mean = (sum as f64) / (count as f64);
            let variance = (sum_squares / count as f64) - (mean * mean);
            if variance < 0.0 {  // Handle floating point errors
                0
            } else {
                (variance.sqrt()).round() as i32
            }
        },
        _ => {
            *error = 1;
            0
        }
    };
    
    // Cache with minimal dependency info to save memory
    let mut minimal_deps = HashSet::new();
    minimal_deps.insert((start_row, start_col));
    minimal_deps.insert((start_row, end_col));
    minimal_deps.insert((end_row, start_col));
    minimal_deps.insert((end_row, end_col));
    
    RANGE_CACHE.with(|cache| {
        cache.borrow_mut().insert(cache_key.to_string(), (result, minimal_deps));
    });
    
    result
}

fn parse_expr<'a>(sheet: &CloneableSheet<'a>, input: &mut &str, cur_row: i32, cur_col: i32, error: &mut i32) -> i32 {
    let mut result = parse_term(sheet, input, cur_row, cur_col, error);
    if *error != 0 { return 0; }
    skip_spaces(input);
    while input.starts_with('+') || input.starts_with('-') {
        let op = input.chars().next().unwrap();
        *input = &input[1..];
        skip_spaces(input);
        let term_value = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 { return 0; }
        if op == '+' { result += term_value; } else { result -= term_value; }
        skip_spaces(input);
    }
    skip_spaces(input);
    if !input.is_empty() && !input.starts_with(')') {
        if !input.chars().all(|ch| ch.is_whitespace()) {
            *error = 1;
        }
    }
    result
}

fn parse_term<'a>(sheet: &CloneableSheet<'a>, input: &mut &str, cur_row: i32, cur_col: i32, error: &mut i32) -> i32 {
    let mut value = parse_factor(sheet, input, cur_row, cur_col, error);
    if *error != 0 { return 0; }
    skip_spaces(input);
    while input.starts_with('*') || input.starts_with('/') {
        let op = input.chars().next().unwrap();
        *input = &input[1..];
        skip_spaces(input);
        let factor_value = parse_factor(sheet, input, cur_row, cur_col, error);
        if *error != 0 { return 0; }
        if op == '/' {
            if factor_value == 0 { *error = 3; return 0; }
            value /= factor_value;
        } else {
            value *= factor_value;
        }
        skip_spaces(input);
    }
    value
}

fn parse_factor<'a>(sheet: &CloneableSheet<'a>, input: &mut &str, cur_row: i32, cur_col: i32, error: &mut i32) -> i32 {
    skip_spaces(input);
    if input.is_empty() {
        *error = 1;
        return 0;
    }
    let ch = input.chars().next().unwrap();
    if ch.is_alphabetic() {
        // Read token (could be function or cell reference).
        let mut token = String::new();
        while let Some(ch) = input.chars().next() {
            if ch.is_alphabetic() {
                token.push(ch);
                *input = &input[ch.len_utf8()..];
            } else {
                break;
            }
        }
        skip_spaces(input);
        if input.starts_with('(') {
            *input = &input[1..]; // Skip '('
            skip_spaces(input);
            if token == "SLEEP" {
                let sleep_time = parse_expr(sheet, input, cur_row, cur_col, error);
                if *error != 0 { return 0; }
                skip_spaces(input);
                if input.starts_with(')') { *input = &input[1..]; }
                if sleep_time < 0 {
                    return sleep_time;
                } else {
                    sleep(Duration::from_secs(sleep_time as u64));
                    return sleep_time;
                }
            } else if token == "MIN" || token == "MAX" || token == "SUM" ||
                      token == "AVG" || token == "STDEV" {
                let close_paren = input.find(')').unwrap_or(input.len());
                let range_str = &input[..close_paren];
                let val = evaluate_range_function(sheet, &token, range_str, error);
                *input = if close_paren < input.len() { &input[close_paren+1..] } else { "" };
                return val;
            } else {
                // Unknown function: skip until ')'
                if let Some(pos) = input.find(')') {
                    *input = &input[pos+1..];
                } else {
                    *error = 1;
                }
                return 0;
            }
        } else {
            // Not a function call; treat token as a cell reference.
            // After reading the alphabetic token, also read the following digits.
            let mut cell_ref = token;
            while let Some(ch) = input.chars().next() {
                if ch.is_digit(10) {
                    cell_ref.push(ch);
                    *input = &input[ch.len_utf8()..];
                } else {
                    break;
                }
            }
            if let Some((r, c)) = crate::sheet::cell_name_to_coords(&cell_ref) {
                if r < 0 || r >= sheet.total_rows() || c < 0 || c >= sheet.total_cols() {
                    *error = 4;
                    return 0;
                }
                if let Some(cell) = sheet.get_cell(r, c) {
                    if cell.status == CellStatus::Error {
                        *error = 3;
                        return 0;
                    }
                    return cell.value;
                } else {
                    *error = 4;
                    return 0;
                }
            } else {
                *error = 1;
                return 0;
            }
        }
    }
    if ch.is_digit(10) || (ch == '-' && input.chars().nth(1).map(|c| c.is_digit(10)).unwrap_or(false)) {
        let mut sign = 1;
        if input.starts_with('-') { sign = -1; *input = &input[1..]; }
        let mut number = 0;
        while let Some(ch) = input.chars().next() {
            if ch.is_digit(10) {
                number = number * 10 + ch.to_digit(10).unwrap() as i32;
                *input = &input[ch.len_utf8()..];
            } else { break; }
        }
        return sign * number;
    }
    if ch == '(' {
        *input = &input[1..];
        let val = parse_expr(sheet, input, cur_row, cur_col, error);
        if *error != 0 { return 0; }
        if input.starts_with(')') { *input = &input[1..]; }
        return val;
    }
    *error = 1;
    0
}

// New function to build and evaluate AST
fn evaluate_ast<'a>(sheet: &CloneableSheet<'a>, ast: &ASTNode, cur_row: i32, cur_col: i32, error: &mut i32) -> i32 {
    match ast {
        ASTNode::Literal(val) => *val,
        ASTNode::CellRef(row, col) => {
            if *row < 0 || *row >= sheet.total_rows() || *col < 0 || *col >= sheet.total_cols() {
                *error = 4;
                return 0;
            }
            if let Some(cell) = sheet.get_cell(*row, *col) {
                if cell.status == CellStatus::Error {
                    *error = 3;
                    return 0;
                }
                cell.value
            } else {
                *error = 4;
                0
            }
        },
        ASTNode::BinaryOp(op, left, right) => {
            let left_val = evaluate_ast(sheet, left, cur_row, cur_col, error);
            if *error != 0 { return 0; }
            
            let right_val = evaluate_ast(sheet, right, cur_row, cur_col, error);
            if *error != 0 { return 0; }
            
            match op {
                '+' => left_val + right_val,
                '-' => left_val - right_val,
                '*' => left_val * right_val,
                '/' => {
                    if right_val == 0 {
                        *error = 3;
                        return 0;
                    }
                    left_val / right_val
                },
                _ => {
                    *error = 1;
                    0
                }
            }
        },
        ASTNode::RangeFunction(func_name, range_str) => {
            evaluate_range_function(sheet, func_name, range_str, error)
        },
        ASTNode::SleepFunction(duration) => {
            let sleep_time = evaluate_ast(sheet, duration, cur_row, cur_col, error);
            if *error != 0 { return 0; }
            
            if sleep_time < 0 {
                return sleep_time;
            } else {
                sleep(Duration::from_secs(sleep_time as u64));
                return sleep_time;
            }
        }
    }
}

/// Public API: evaluate_formula 
pub fn evaluate_formula<'a>(
    sheet: &CloneableSheet<'a>,
    formula: &str,
    current_row: i32,
    current_col: i32,
    error: &mut i32,
    status_msg: &mut String,
) -> i32 {
    let trimmed = formula.trim().to_string();
    if trimmed.is_empty() {
        *error = 1;
        status_msg.clear();
        status_msg.push_str("Memory allocation error");
        return 0;
    }
    let mut input = trimmed.as_str();
    *error = 0;
    let result = parse_expr(sheet, &mut input, current_row, current_col, error);
    if *error == 1 {
        status_msg.clear();
        status_msg.push_str("Invalid formula");
        return 0;
    } else if *error == 2 {
        status_msg.clear();
        status_msg.push_str("Invalid range");
        return 0;
    } else if *error == 3 {
        return 0;
    }
    result
}

// Function to clear the thread-local cache
pub fn clear_range_cache() {
    RANGE_CACHE.with(|cache| {
        cache.borrow_mut().clear();
    });
}

// Add a function to invalidate cache entries for a specific cell
pub fn invalidate_cache_for_cell(row: i32, col: i32) {
    RANGE_CACHE.with(|cache| {
        let mut cache_ref = cache.borrow_mut();
        
        // Find all cache entries that include this cell in their dependencies
        let keys_to_remove: Vec<String> = cache_ref.iter()
            .filter(|(_, (_, deps))| deps.contains(&(row, col)))
            .map(|(key, _)| key.clone())
            .collect();
        
        // Remove those entries
        for key in keys_to_remove {
            cache_ref.remove(&key);
        }
    });
}
