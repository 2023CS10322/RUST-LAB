#![allow(warnings)]
use crate::sheet::cell_name_to_coords;
use crate::sheet::{CachedRange, CellStatus, CloneableSheet, Spreadsheet};
use std::collections::{HashMap, HashSet};
use std::thread::sleep;
use std::time::Duration;

// Define the AST node enum for formula parsing
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl Value {
    pub fn as_number(&self) -> Option<f64> {
        if let Value::Number(n) = self {
            Some(*n)
        } else {
            None
        }
    }
    pub fn as_bool(&self) -> Option<bool> {
        if let Value::Bool(b) = self {
            Some(*b)
        } else {
            None
        }
    }
    pub fn as_text(&self) -> Option<&str> {
        if let Value::Text(s) = self {
            Some(s)
        } else {
            None
        }
    }
}

#[derive(Clone, Debug)]
pub enum ASTNode {
    Literal(i32),
    CellRef(i32, i32),
    BinaryOp(char, Box<ASTNode>, Box<ASTNode>),
    RangeFunction(String, String), // Function name and range string
    SleepFunction(Box<ASTNode>),
}

// Keep the cache in thread_local storage for thread safety
thread_local! {
    pub static RANGE_CACHE: std::cell::RefCell<HashMap<String, (i32, HashSet<(i32, i32)>)>> =
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

pub fn evaluate_range_function<'a>(
    sheet: &CloneableSheet<'a>,
    func_name: &str,
    range_str: &str,
    error: &mut i32,
) -> i32 {
    // Check if we have this range cached
    let cache_key = format!("{}({})", func_name, range_str);

    // Try to get from thread-local cache with improved validation
    if let Some((cached_value, _)) = RANGE_CACHE.with(|cache| {
        cache
            .borrow()
            .get(&cache_key)
            .map(|(val, deps)| (*val, deps.clone()))
    }) {
        return cached_value;
    }

    if let Some(colon_pos) = range_str.find(':') {
        let cell1 = range_str[..colon_pos].trim();
        let cell2 = range_str[colon_pos + 1..].trim();
        let (start_row, start_col) = match crate::sheet::cell_name_to_coords(cell1) {
            Some(coords) => coords,
            None => {
                *error = 1;
                return 0;
            }
        };
        let (end_row, end_col) = match crate::sheet::cell_name_to_coords(cell2) {
            Some(coords) => coords,
            None => {
                *error = 1;
                return 0;
            }
        };
        if start_row > end_row || start_col > end_col {
            *error = 2;
            return 0;
        }

        // Check bounds
        if start_row < 0
            || end_row >= sheet.total_rows()
            || start_col < 0
            || end_col >= sheet.total_cols()
        {
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
                    if value < min_val {
                        min_val = value;
                    }
                    if value > max_val {
                        max_val = value;
                    }
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
            }
            _ => {
                *error = 1;
                0
            }
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
pub fn evaluate_large_range<'a>(
    sheet: &CloneableSheet<'a>,
    func_name: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    error: &mut i32,
    cache_key: &str,
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

                        if value < min_val {
                            min_val = value;
                        }
                        if value > max_val {
                            max_val = value;
                        }
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
                *error = 3; // Overflow
                return 0;
            }
            sum as i32
        }
        "AVG" => {
            let avg = sum / (count as i64);
            if avg > i32::MAX as i64 || avg < i32::MIN as i64 {
                *error = 3; // Overflow
                return 0;
            }
            avg as i32
        }
        "STDEV" => {
            let mean = (sum as f64) / (count as f64);
            let variance = (sum_squares / count as f64) - (mean * mean);
            if variance < 0.0 {
                // Handle floating point errors
                0
            } else {
                (variance.sqrt()).round() as i32
            }
        }
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
        cache
            .borrow_mut()
            .insert(cache_key.to_string(), (result, minimal_deps));
    });

    result
}

pub fn parse_expr<'a>(
    sheet: &CloneableSheet<'a>,
    input: &mut &str,
    cur_row: i32,
    cur_col: i32,
    error: &mut i32,
) -> i32 {
    // 1) Parse the initial term.
    let mut value = parse_term(sheet, input, cur_row, cur_col, error);
    if *error != 0 {
        return 0;
    }
    skip_spaces(input);

    // 2) Optional comparison operators.
    if input.starts_with(">=") {
        *input = &input[2..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        value = if value >= rhs { 1 } else { 0 };
        skip_spaces(input);
    } else if input.starts_with(">") {
        *input = &input[1..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        value = if value > rhs { 1 } else { 0 };
        skip_spaces(input);
    } else if input.starts_with("<=") {
        *input = &input[2..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        value = if value <= rhs { 1 } else { 0 };
        skip_spaces(input);
    } else if input.starts_with("<") {
        *input = &input[1..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        value = if value < rhs { 1 } else { 0 };
        skip_spaces(input);
    } else if input.starts_with("==") {
        *input = &input[2..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        value = if value == rhs { 1 } else { 0 };
        skip_spaces(input);
    }

    // 3) Then handle addition and subtraction.
    while let Some(op) = input.chars().next() {
        if op != '+' && op != '-' {
            break;
        }
        *input = &input[1..];
        skip_spaces(input);
        let rhs = parse_term(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        if op == '+' {
            value += rhs
        } else {
            value -= rhs
        }
        skip_spaces(input);
    }

    // 4) Finally, allow ')' or ',' (for IF) or whitespace/end without error.
    skip_spaces(input);
    if !input.is_empty() {
        match input.chars().next().unwrap() {
            ')' | ',' => { /* OK */ }
            ch if ch.is_whitespace() => { /* OK */ }
            _ => *error = 1,
        }
    }

    value
}

pub fn parse_term<'a>(
    sheet: &CloneableSheet<'a>,
    input: &mut &str,
    cur_row: i32,
    cur_col: i32,
    error: &mut i32,
) -> i32 {
    let mut value = parse_factor(sheet, input, cur_row, cur_col, error);
    if *error != 0 {
        return 0;
    }
    skip_spaces(input);
    while input.starts_with('*') || input.starts_with('/') {
        let op = input.chars().next().unwrap();
        *input = &input[1..];
        skip_spaces(input);
        let factor_value = parse_factor(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        if op == '/' {
            if factor_value == 0 {
                *error = 3;
                return 0;
            }
            value /= factor_value;
        } else {
            value *= factor_value;
        }
        skip_spaces(input);
    }
    value
}

fn parse_range_bounds(s: &str, error: &mut i32) -> Option<(i32, i32, i32, i32)> {
    if let Some(colon) = s.find(':') {
        let a = &s[..colon];
        let b = &s[colon + 1..];
        if let (Some((r1, c1)), Some((r2, c2))) = (cell_name_to_coords(a), cell_name_to_coords(b)) {
            return Some((r1, c1, r2, c2));
        }
    }
    *error = 1;
    None
}

pub fn parse_factor<'a>(
    sheet: &CloneableSheet<'a>,
    input: &mut &str,
    cur_row: i32,
    cur_col: i32,
    error: &mut i32,
) -> i32 {
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

            if token == "IF" && cfg!(feature = "advanced_formulas") {
                let cond = parse_expr(sheet, input, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                skip_spaces(input);
                if !input.starts_with(',') {
                    *error = 1;
                    return 0;
                }
                *input = &input[1..];
                skip_spaces(input);

                let tv = parse_expr(sheet, input, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                skip_spaces(input);
                if !input.starts_with(',') {
                    *error = 1;
                    return 0;
                }
                *input = &input[1..];
                skip_spaces(input);

                let fv = parse_expr(sheet, input, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                skip_spaces(input);
                if input.starts_with(')') {
                    *input = &input[1..];
                }

                return if cond != 0 { tv } else { fv };
            }
            // COUNTIF(range, condition)
            else if token == "COUNTIF" && cfg!(feature = "advanced_formulas") {
                let close = input.find(')').unwrap_or(input.len());
                // extract the raw args string, then advance input
                let args = &input[..close];
                *input = &input[close..];

                // split into range and criterion
                let parts: Vec<&str> = args.splitn(2, ',').map(str::trim).collect();
                if parts.len() != 2 {
                    *error = 1;
                    return 0;
                }

                // parse the range bounds A1:B2
                let (r1, c1, r2, c2) = match parse_range_bounds(parts[0], error) {
                    Some(b) => b,
                    None => return 0,
                };

                let mut count = 0;
                // decide if criterion is a quoted comparison or a simple numeric equality
                let crit = parts[1];
                let (op, threshold) = if crit.starts_with('"') && crit.ends_with('"') {
                    // strip quotes
                    let inner = &crit[1..crit.len() - 1];
                    // find operator prefix
                    let ops = [">=", "<=", "<>", ">", "<", "="]; // <> for not equal
                    let mut found = None;
                    for &candidate in &ops {
                        if inner.starts_with(candidate) {
                            if let Ok(val) = inner[candidate.len()..].trim().parse::<i32>() {
                                found = Some((candidate, val));
                            }
                            break;
                        }
                    }
                    match found {
                        Some(f) => f,
                        None => {
                            *error = 1;
                            return 0;
                        }
                    }
                } else {
                    // default: numeric equality
                    // parse once
                    let mut crit_s = crit;
                    let val = parse_expr(sheet, &mut crit_s, cur_row, cur_col, error);
                    if *error != 0 {
                        return 0;
                    }
                    // treat as "=val"
                    ("=", val)
                };

                // iterate cells
                for rr in r1..=r2 {
                    for cc in c1..=c2 {
                        if let Some(cell) = sheet.get_cell(rr, cc) {
                            if cell.status == CellStatus::Error {
                                *error = 3;
                                return 0;
                            }
                            let v = cell.value;
                            let m = match op {
                                ">" => v > threshold,
                                ">=" => v >= threshold,
                                "<" => v < threshold,
                                "<=" => v <= threshold,
                                "=" => v == threshold,
                                "<>" => v != threshold,
                                _ => false,
                            };
                            if m {
                                count += 1;
                            }
                        }
                    }
                }
                if input.starts_with(')') {
                    *input = &input[1..];
                }
                return count;
            }
            // SUMIF(range, condition, sum_range)
            // SUMIF(range, criterion, sum_range)
            // Inside parse_factor, after matching token == "SUMIF":
            else if token == "SUMIF" && cfg!(feature = "advanced_formulas") {
                // Grab everything up to the closing ')'
                let close = input.find(')').unwrap_or(input.len());
                let args = &input[..close];
                *input = &input[close..];

                // Split into exactly three comma‑separated parts
                let parts: Vec<&str> = args.splitn(3, ',').map(str::trim).collect();
                if parts.len() != 3 {
                    *error = 1;
                    return 0;
                }

                // 1) parse the test range A1:B2 → (r1,c1,r2,c2)
                let (r1, c1, r2, c2) = match parse_range_bounds(parts[0], error) {
                    Some(b) => b,
                    None => return 0,
                };
                // 2) parse the sum range  C1:D2 → (s1,t1,s2,t2)
                let (s1, t1, s2, t2) = match parse_range_bounds(parts[2], error) {
                    Some(b) => b,
                    None => return 0,
                };

                // ── REQUIRE IDENTICAL DIMENSIONS ──
                let rows_test = r2 - r1;
                let cols_test = c2 - c1;
                let rows_sum = s2 - s1;
                let cols_sum = t2 - t1;
                if rows_test != rows_sum || cols_test != cols_sum {
                    *error = 1;
                    return 0;
                }

                // 3) parse the criterion, either quoted >5 style or plain numeric
                let crit = parts[1];
                let (op, threshold) = if crit.starts_with('\"') && crit.ends_with('\"') {
                    // strip the quotes and detect operator
                    let inner = &crit[1..crit.len() - 1];
                    let ops = [">=", "<=", "<>", ">", "<", "="];
                    let mut found = None;
                    for &candidate in &ops {
                        if inner.starts_with(candidate) {
                            if let Ok(val) = inner[candidate.len()..].trim().parse::<i32>() {
                                found = Some((candidate, val));
                            }
                            break;
                        }
                    }
                    match found {
                        Some(f) => f,
                        None => {
                            *error = 1;
                            return 0;
                        }
                    }
                } else {
                    // one‑time numeric equality
                    let mut crit_s = crit;
                    let val = parse_expr(sheet, &mut crit_s, cur_row, cur_col, error);
                    if *error != 0 {
                        return 0;
                    }
                    ("=", val)
                };

                // 4) loop over every cell in the test range and sum matching cells
                let mut total = 0;
                for dr in 0..=rows_test {
                    for dc in 0..=cols_test {
                        let rr = r1 + dr;
                        let cc = c1 + dc;
                        if let Some(cell) = sheet.get_cell(rr, cc) {
                            if cell.status == CellStatus::Error {
                                *error = 3;
                                return 0;
                            }
                            let v = cell.value;
                            let keep = match op {
                                ">" => v > threshold,
                                ">=" => v >= threshold,
                                "<" => v < threshold,
                                "<=" => v <= threshold,
                                "=" => v == threshold,
                                "<>" => v != threshold,
                                _ => false,
                            };
                            if keep {
                                // same offset into sum_range
                                let sr = s1 + dr;
                                let sc = t1 + dc;
                                if let Some(sumc) = sheet.get_cell(sr, sc) {
                                    if sumc.status == CellStatus::Error {
                                        *error = 3;
                                        return 0;
                                    }
                                    total += sumc.value;
                                }
                            }
                        }
                    }
                }

                // consume the closing ')'
                if input.starts_with(')') {
                    *input = &input[1..];
                }
                return total;
            }
            // ROUND(value, digits)
            else if token == "ROUND" && cfg!(feature = "advanced_formulas") {
                let close = input.find(')').unwrap_or(input.len());
                let args = &input[..close];
                *input = &input[close..];
                let parts: Vec<&str> = args.splitn(2, ',').map(str::trim).collect();
                if parts.len() != 2 {
                    *error = 1;
                    return 0;
                }
                let mut s0 = parts[0];
                let mut s1 = parts[1];
                let val = parse_expr(sheet, &mut s0, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                let digs = parse_expr(sheet, &mut s1, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                // NEW: drop last 'digs' digits
                let factor = 10_i32.pow(digs as u32);
                let truncated = val / factor;
                if input.starts_with(')') {
                    *input = &input[1..];
                }
                return truncated;
            } else if token == "SLEEP" {
                let sleep_time = parse_expr(sheet, input, cur_row, cur_col, error);
                if *error != 0 {
                    return 0;
                }
                skip_spaces(input);
                if input.starts_with(')') {
                    *input = &input[1..];
                }
                if sleep_time < 0 {
                    return sleep_time;
                } else {
                    sleep(Duration::from_secs(sleep_time as u64));
                    return sleep_time;
                }
            } else if token == "MIN"
                || token == "MAX"
                || token == "SUM"
                || token == "AVG"
                || token == "STDEV"
            {
                let close_paren = input.find(')').unwrap_or(input.len());
                let range_str = &input[..close_paren];
                let val = evaluate_range_function(sheet, &token, range_str, error);
                *input = if close_paren < input.len() {
                    &input[close_paren + 1..]
                } else {
                    ""
                };
                return val;
            } else {
                // Unknown function: skip until ')'
                if let Some(pos) = input.find(')') {
                    *input = &input[pos + 1..];
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
    if ch.is_digit(10)
        || (ch == '-'
            && input
                .chars()
                .nth(1)
                .map(|c| c.is_digit(10))
                .unwrap_or(false))
    {
        let mut sign = 1;
        if input.starts_with('-') {
            sign = -1;
            *input = &input[1..];
        }
        let mut number = 0;
        while let Some(ch) = input.chars().next() {
            if ch.is_digit(10) {
                number = number * 10 + ch.to_digit(10).unwrap() as i32;
                *input = &input[ch.len_utf8()..];
            } else {
                break;
            }
        }
        return sign * number;
    }
    if ch == '(' {
        *input = &input[1..];
        let val = parse_expr(sheet, input, cur_row, cur_col, error);
        if *error != 0 {
            return 0;
        }
        if input.starts_with(')') {
            *input = &input[1..];
        }
        return val;
    }
    *error = 1;
    0
}

// New function to build and evaluate AST
pub fn evaluate_ast<'a>(
    sheet: &CloneableSheet<'a>,
    ast: &ASTNode,
    cur_row: i32,
    cur_col: i32,
    error: &mut i32,
) -> i32 {
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
        }
        ASTNode::BinaryOp(op, left, right) => {
            let left_val = evaluate_ast(sheet, left, cur_row, cur_col, error);
            if *error != 0 {
                return 0;
            }

            let right_val = evaluate_ast(sheet, right, cur_row, cur_col, error);
            if *error != 0 {
                return 0;
            }

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
                }
                _ => {
                    *error = 1;
                    0
                }
            }
        }
        ASTNode::RangeFunction(func_name, range_str) => {
            evaluate_range_function(sheet, func_name, range_str, error)
        }
        ASTNode::SleepFunction(duration) => {
            let sleep_time = evaluate_ast(sheet, duration, cur_row, cur_col, error);
            if *error != 0 {
                return 0;
            }

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
        let keys_to_remove: Vec<String> = cache_ref
            .iter()
            .filter(|(_, (_, deps))| deps.contains(&(row, col)))
            .map(|(key, _)| key.clone())
            .collect();

        // Remove those entries
        for key in keys_to_remove {
            cache_ref.remove(&key);
        }
    });
}

// at the bottom of src/parser.rs
// ─── parser.rs ──────────────────────────────────────────────────────────────
// your existing `pub fn evaluate_formula(…) { … }` etc.
// ────────────────────────────────────────────────────────────────────────────

#[cfg(test)]


mod tests {
    use super::*;  // brings in evaluate_formula, clear_range_cache, ASTNode, parse, evaluate_ast, etc.

    
    /// Does the very basics: literals and + - * /, with and without parens.
    #[test]
    fn basic_arithmetic_and_parens() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        // literal
        assert_eq!(evaluate_formula(&cs, "42", 0, 0, &mut err, &mut status), 42);
        assert_eq!(err, 0);

        // ops
        assert_eq!(evaluate_formula(&cs, "2+3*4", 0, 0, &mut err, &mut status), 14);
        assert_eq!(evaluate_formula(&cs, "(2+3)*4", 0, 0, &mut err, &mut status), 20);

        // minus‐unary
        assert_eq!(evaluate_formula(&cs, "-5+10", 0, 0, &mut err, &mut status), 5);
    }

    /// Division by zero should set error code 3
    #[test]
    fn division_by_zero_sets_error() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        evaluate_formula(&cs, "10/0", 0, 0, &mut err, &mut status);
        assert_eq!(err, 3);

        err = 0;
        evaluate_formula(&cs, "10/(5-5)", 0, 0, &mut err, &mut status);
        assert_eq!(err, 3);
    }

    /// Cell references pick up values from the sheet
    #[test]
    fn cell_references_and_sum_functions() {
        let mut sheet = Spreadsheet::new(2, 2);
        sheet.update_cell_value(0, 0, 7, CellStatus::Ok);  // A1 = 7
        sheet.update_cell_value(0, 1, 3, CellStatus::Ok);  // B1 = 3
        sheet.update_cell_value(1, 0, 2, CellStatus::Ok);  // A2 = 2
        sheet.update_cell_value(1, 1, 4, CellStatus::Ok);  // B2 = 4

        let cs = CloneableSheet::new(&sheet);
        let mut err = 0; let mut status = String::new();

        // single‐cell refs
        assert_eq!(evaluate_formula(&cs, "A1", 0, 0, &mut err, &mut status), 7);
        assert_eq!(evaluate_formula(&cs, "B2", 0, 0, &mut err, &mut status), 4);

        // range functions
        assert_eq!(evaluate_formula(&cs, "SUM(A1:B2)", 0, 0, &mut err, &mut status), 16);
        assert_eq!(evaluate_formula(&cs, "MIN(A1:B2)", 0, 0, &mut err, &mut status), 2);
        assert_eq!(evaluate_formula(&cs, "MAX(A1:B2)", 0, 0, &mut err, &mut status), 7);
        assert_eq!(evaluate_formula(&cs, "AVG(A1:B2)", 0, 0, &mut err, &mut status), 4);
    }

    /// Whitespace-only or empty formulas are errors
    #[test]
    fn empty_and_whitespace_formulas_error() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0; let mut status = String::new();

        evaluate_formula(&cs, "", 0, 0, &mut err, &mut status);
        assert_eq!(err, 1);

        err = 0; status.clear();
        evaluate_formula(&cs, "   ", 0, 0, &mut err, &mut status);
        assert_eq!(err, 1);
    }

    /// Caching works: clear_range_cache forces a recompute
    #[test]
    fn cache_cleared_recomputes_sum() {
        let mut sheet = Spreadsheet::new(1, 2);
        sheet.update_cell_value(0, 0, 5, CellStatus::Ok);
        sheet.update_cell_value(0, 1, 6, CellStatus::Ok);

        let cs1 = CloneableSheet::new(&sheet);
        let mut err = 0; let mut status = String::new();
        let first = evaluate_formula(&cs1, "SUM(A1:B1)", 0, 0, &mut err, &mut status);

        // change underlying value & clear cache
        sheet.update_cell_value(0, 0, 8, CellStatus::Ok);
        clear_range_cache();

        let cs2 = CloneableSheet::new(&sheet);
        let second = evaluate_formula(&cs2, "SUM(A1:B1)", 0, 0, &mut err, &mut status);

        assert_ne!(first, second);
        assert_eq!(second, 14);
    }
    // at the bottom of src/parser.rs


    /// Helper: build a cloneable sheet with a few (row, col, value) tuples
   

// In tests/sheet_tests.rs (or wherever your sheet unit tests live)



    #[test]
    fn test_update_cell_value_and_status() {
        let mut sheet = Spreadsheet::new(2, 2);
        // initially all cells are zero/Ok
        assert_eq!(sheet.get_cell_value(0, 0), 0);
        assert_eq!(sheet.get_cell_status(0, 0), CellStatus::Ok);

        // update A1 to 42 with Error status
        sheet.update_cell_value(0, 0, 42, CellStatus::Error);
        assert_eq!(sheet.get_cell_value(0, 0), 42);
        assert_eq!(sheet.get_cell_status(0, 0), CellStatus::Error);

        // update B2 to -7 with Ok status
        sheet.update_cell_value(1, 1, -7, CellStatus::Ok);
        assert_eq!(sheet.get_cell_value(1, 1), -7);
        assert_eq!(sheet.get_cell_status(1, 1), CellStatus::Ok);
    }

    #[test]
    fn test_parser_observes_updated_values() {
        let mut sheet = Spreadsheet::new(2, 2);

        // set A1=5, B1=7
        sheet.update_cell_value(0, 0, 5, CellStatus::Ok);
        sheet.update_cell_value(0, 1, 7, CellStatus::Ok);

        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut msg = String::new();

        // SUM(A1:B1) == 12
        let s1 = evaluate_formula(&cs, "SUM(A1:B1)", 0, 0, &mut err, &mut msg);
        assert_eq!(s1, 12);
        assert_eq!(err, 0);

        // change A1 to 10, clear cache, re-sum → 17
        sheet.update_cell_value(0, 0, 10, CellStatus::Ok);
        clear_range_cache();

        let cs2 = CloneableSheet::new(&sheet);
        let s2 = evaluate_formula(&cs2, "SUM(A1:B1)", 0, 0, &mut err, &mut msg);
        assert_eq!(s2, 17);
        assert_eq!(err, 0);
    }

// In tests/parser_tests.rs



   

   


    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn advanced_if_countif_sumif() {
        let cs = sheet_with(&[
            (0, 0, 10),
            (0, 1, 20),
            (1, 0, 30),
            (1, 1, 40),
        ]);
        let mut err = 0;
        let mut status = String::new();

        // IF()
        assert_eq!(
            evaluate_formula(&cs, "IF(A1<B1,1,0)", 0, 0, &mut err, &mut status),
            1
        );
        // COUNTIF
        assert_eq!(
            evaluate_formula(&cs, "COUNTIF(A1:B2,\">25\")", 0, 0, &mut err, &mut status),
            2
        );
        // SUMIF
        assert_eq!(
            evaluate_formula(
                &cs,
                "SUMIF(A1:B2,\">25\",A1:B2)",
                0,
                0,
                &mut err,
                &mut status
            ),
            70
        );
    }


    #[test]
    fn test_number_and_basic_ops() {
        // 1. Own the sheet
        let sheet = Spreadsheet::new(1, 1);
        // 2. Borrow it for parsing
        let cs = CloneableSheet::new(&sheet);
        // 3. Run your formula tests
        let mut err = 0;
        let mut status = String::new();
        assert_eq!(evaluate_formula(&cs, "42", 0, 0, &mut err, &mut status), 42);
        assert_eq!(evaluate_formula(&cs, "2+3*4-5", 0, 0, &mut err, &mut status), 2 + 3*4 - 5);
        assert_eq!(err, 0);
    }

    #[test]
    fn test_parens_and_unary() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);

        let mut err = 0;
        let mut status = String::new();
        let result = evaluate_formula(&cs, "-(1+2)*3", 0, 0, &mut err, &mut status);

        // Parser does not support unary minus before parentheses, so it should error:
        assert_eq!(result, 0, "Expected 0 when unary- grouping is unsupported");
        assert_eq!(err, 1, "Expected error code 1 for invalid formula");
    }


    #[test]
    fn test_cell_refs_and_sum() {
        // Build & populate a 2×2 sheet
        let mut sheet = Spreadsheet::new(2, 2);
        sheet.update_cell_value(0, 0, 10, CellStatus::Ok); // A1
        sheet.update_cell_value(0, 1, 20, CellStatus::Ok); // B1
        sheet.update_cell_value(1, 0,  5, CellStatus::Ok); // A2

        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "A1",          0, 0, &mut err, &mut status), 10);
        assert_eq!(evaluate_formula(&cs, "SUM(A1:B2)", 0, 0, &mut err, &mut status), 10+20+5+0);
        assert_eq!(err, 0);
    }

    #[test]
    fn test_invalid_and_errors() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);

        let mut err = 0;
        let mut status = String::new();

        // empty formula
        evaluate_formula(&cs, "", 0, 0, &mut err, &mut status);
        assert_eq!(err, 1);

        // bad range syntax
        err = 0; status.clear();
        evaluate_formula(&cs, "SUM(A1B2)", 0, 0, &mut err, &mut status);
        assert_eq!(err, 1);

        // divide by zero
        err = 0; status.clear();
        evaluate_formula(&cs, "1/0", 0, 0, &mut err, &mut status);
        assert_eq!(err, 3);
    }

    #[test]
    fn test_cache_and_clear() {
        // 1) Build sheet and initial sum
        let mut sheet = Spreadsheet::new(1, 2);
        sheet.update_cell_value(0, 0, 2, CellStatus::Ok);
        sheet.update_cell_value(0, 1, 3, CellStatus::Ok);
        let cs1 = CloneableSheet::new(&sheet);

        let mut err = 0;
        let mut status = String::new();
        let v1 = evaluate_formula(&cs1, "SUM(A1:B1)", 0, 0, &mut err, &mut status);
        assert_eq!(v1, 5);

        // 2) Mutate, clear the parser cache, re-evaluate
        sheet.update_cell_value(0, 0, 7, CellStatus::Ok);
        clear_range_cache();
        let cs2 = CloneableSheet::new(&sheet);
        let v2 = evaluate_formula(&cs2, "SUM(A1:B1)", 0, 0, &mut err, &mut status);
        assert_eq!(v2, 7 + 3);
    }





// ─── tests for parser.rs ──────────────────────────────────────────────────────
#[cfg(test)]
mod parser_tests {
    use super::*;
    use crate::sheet::{Spreadsheet, CloneableSheet, CellStatus};

    #[test]
    fn value_enum_helpers() {
        assert_eq!(Value::Number(3.14).as_number(), Some(3.14));
        assert_eq!(Value::Bool(true).as_bool(), Some(true));
        assert_eq!(Value::Text("hello".into()).as_text(), Some("hello"));
        // mismatch
        assert_eq!(Value::Number(5.0).as_bool(), None);
    }

    #[test]
    fn basic_arithmetic_and_parens() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "2+3*4", 0, 0, &mut err, &mut status), 14);
        assert_eq!(evaluate_formula(&cs, "(2+3)*4", 0, 0, &mut err, &mut status), 20);
    }

    #[test]
    fn comparison_operators() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "2>1", 0, 0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "2<1", 0, 0, &mut err, &mut status), 0);
        assert_eq!(evaluate_formula(&cs, "2==2", 0, 0, &mut err, &mut status), 1);
    }

    #[test]
    fn range_functions_min_max_avg_stdev() {
        let mut s = Spreadsheet::new(3, 1);
        s.update_cell_value(0, 0, 1, CellStatus::Ok);
        s.update_cell_value(1, 0, 3, CellStatus::Ok);
        s.update_cell_value(2, 0, 5, CellStatus::Ok);

        let cs = CloneableSheet::new(&s);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "MIN(A1:A3)", 0, 0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "MAX(A1:A3)", 0, 0, &mut err, &mut status), 5);
        assert_eq!(evaluate_formula(&cs, "SUM(A1:A3)", 0, 0, &mut err, &mut status), 9);
        assert_eq!(evaluate_formula(&cs, "AVG(A1:A3)", 0, 0, &mut err, &mut status), 3);
        // Variance = ((1−3)² + (3−3)² + (5−3)²)/3 = (4+0+4)/3≈2.666→√≈1.63→round→2
        assert_eq!(
            evaluate_formula(&cs, "STDEV(A1:A3)", 0, 0, &mut err, &mut status),
            2
        );
    }

    #[test]
    fn parse_unknown_and_error_cases() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        // unknown function → returns 0
        let v1 = evaluate_formula(&cs, "FOO(1)", 0, 0, &mut err, &mut status);
        assert_eq!(v1, 0);

        // invalid formula → err=1
        err = 0;
        evaluate_formula(&cs, "1?/2", 0, 0, &mut err, &mut status);
        assert_eq!(err, 1);
    }

    #[test]
    fn test_sleep_negative_fast() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        // negative sleep returns immediately
        let v = evaluate_formula(&cs, "SLEEP(-2)", 0, 0, &mut err, &mut status);
        assert_eq!(v, -2);
        assert_eq!(err, 0);
    }
    
    use super::*;
    // 2) grab the sheet types you need for constructing CloneableSheet



    #[test]
    fn comparison_and_arithmetic() {
        let sheet = Spreadsheet::new(1,1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "2>1", 0,0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "2<1", 0,0, &mut err, &mut status), 0);
        assert_eq!(evaluate_formula(&cs, "2==2",0,0, &mut err, &mut status), 1);

        assert_eq!(evaluate_formula(&cs, "3+4*2",0,0, &mut err, &mut status), 11);
        assert_eq!(evaluate_formula(&cs, "(3+4)*2",0,0, &mut err, &mut status), 14);
    }

    #[test]
    fn range_functions_and_errors() {
        let mut s = Spreadsheet::new(3,1);
        s.update_cell_value(0,0,1,CellStatus::Ok);
        s.update_cell_value(1,0,3,CellStatus::Ok);
        s.update_cell_value(2,0,5,CellStatus::Ok);
        let cs = CloneableSheet::new(&s);
        let mut err = 0; let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "SUM(A1:A3)",   0,0, &mut err, &mut status), 9);
        assert_eq!(evaluate_formula(&cs, "MIN(A1:A3)",   0,0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "MAX(A1:A3)",   0,0, &mut err, &mut status), 5);
        assert_eq!(evaluate_formula(&cs, "AVG(A1:A3)",   0,0, &mut err, &mut status), 3);
        assert_eq!(evaluate_formula(&cs, "STDEV(A1:A3)", 0,0, &mut err, &mut status), 2);

        // reversed range → error=2, message="Invalid range"
        err = 0; status.clear();
        assert_eq!(evaluate_formula(&cs, "SUM(B2:A1)", 0,0, &mut err, &mut status), 0);
        assert_eq!(err, 2);
        assert_eq!(status, "Invalid range");
    }

    #[test]
    fn ast_and_cache_invalidation() {
        let mut s = Spreadsheet::new(2,2);
        s.update_cell_value(0,0,1,CellStatus::Ok);
        s.update_cell_value(0,1,2,CellStatus::Ok);

        let cs = CloneableSheet::new(&s);
        let mut err = 0; let mut status = String::new();

        // ASTNode eval
        let lit = ASTNode::Literal(7);
        assert_eq!(evaluate_ast(&cs,&lit,0,0,&mut err), 7);

        let cref = ASTNode::CellRef(0,1);
        assert_eq!(evaluate_ast(&cs,&cref,0,0,&mut err), 2);

        let bop = ASTNode::BinaryOp('+',
                    Box::new(ASTNode::Literal(5)),
                    Box::new(ASTNode::Literal(6)));
        assert_eq!(evaluate_ast(&cs,&bop,0,0,&mut err), 11);

        // unknown op → err=1
        let bad = ASTNode::BinaryOp('?',Box::new(ASTNode::Literal(1)),Box::new(ASTNode::Literal(1)));
        err = 0;
        assert_eq!(evaluate_ast(&cs,&bad,0,0,&mut err), 0);
        assert_eq!(err, 1);

        // clear & invalidate cache
        clear_range_cache();
        let _ = evaluate_formula(&cs, "SUM(A1:B1)", 0,0, &mut err, &mut status);
        invalidate_cache_for_cell(0,0);

        s.update_cell_value(0,0,5,CellStatus::Ok);
        let cs2 = CloneableSheet::new(&s);
        let v2 = evaluate_formula(&cs2, "SUM(A1:B1)", 0,0, &mut err, &mut status);
        assert_eq!(v2, 7);
    }

    #[test]
    fn parse_range_bounds_direct() {
        let mut err = 0;
        assert_eq!(parse_range_bounds("A1:B2", &mut err), Some((0,0,1,1)));
        err = 0;
        assert!(parse_range_bounds("NoColon", &mut err).is_none());
        assert_eq!(err, 1);
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    #[cfg(test)]
mod tests {
    // bring all of parser.rs (including private helpers) into scope
    use super::*;
    // bring in what we need from sheet.rs
    use crate::sheet::{Spreadsheet, CloneableSheet, CellStatus};

    #[test]
    fn value_enum_helpers() {
        assert_eq!(Value::Number(3.14).as_number(), Some(3.14));
        assert_eq!(Value::Bool(true).as_bool(), Some(true));
        assert_eq!(Value::Text("hi".into()).as_text(), Some("hi"));
        assert_eq!(Value::Number(0.0).as_bool(), None);
    }

    #[test]
    fn comparison_and_arithmetic() {
        let sheet = Spreadsheet::new(1,1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "2>1", 0,0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "2<1", 0,0, &mut err, &mut status), 0);
        assert_eq!(evaluate_formula(&cs, "2==2",0,0, &mut err, &mut status), 1);

        assert_eq!(evaluate_formula(&cs, "3+4*2",0,0, &mut err, &mut status), 11);
        assert_eq!(evaluate_formula(&cs, "(3+4)*2",0,0, &mut err, &mut status), 14);
    }

    #[test]
    fn range_functions_and_error_messages() {
        let mut s = Spreadsheet::new(3,1);
        s.update_cell_value(0,0,1,CellStatus::Ok);
        s.update_cell_value(1,0,3,CellStatus::Ok);
        s.update_cell_value(2,0,5,CellStatus::Ok);
        let cs = CloneableSheet::new(&s);
        let mut err = 0; let mut status = String::new();

        assert_eq!(evaluate_formula(&cs, "SUM(A1:A3)",   0,0, &mut err, &mut status), 9);
        assert_eq!(evaluate_formula(&cs, "MIN(A1:A3)",   0,0, &mut err, &mut status), 1);
        assert_eq!(evaluate_formula(&cs, "MAX(A1:A3)",   0,0, &mut err, &mut status), 5);
        assert_eq!(evaluate_formula(&cs, "AVG(A1:A3)",   0,0, &mut err, &mut status), 3);
        assert_eq!(evaluate_formula(&cs, "STDEV(A1:A3)", 0,0, &mut err, &mut status), 2);

        // reversed range → error code 2 + “Invalid range”
        err = 0; status.clear();
        let v = evaluate_formula(&cs, "SUM(B2:A1)", 0,0, &mut err, &mut status);
        assert_eq!(v, 0);
        assert_eq!(err, 2);
        assert_eq!(status, "Invalid range");
    }

    #[test]
    fn ast_and_cache_invalidation() {
        let mut s = Spreadsheet::new(2,2);
        s.update_cell_value(0,0,1,CellStatus::Ok);
        s.update_cell_value(0,1,2,CellStatus::Ok);

        let cs = CloneableSheet::new(&s);
        let mut err = 0; let mut status = String::new();

        // ASTNode::Literal, ::CellRef, ::BinaryOp
        let lit = ASTNode::Literal(7);
        assert_eq!(evaluate_ast(&cs, &lit, 0,0, &mut err), 7);

        let cref = ASTNode::CellRef(0,1);
        assert_eq!(evaluate_ast(&cs, &cref, 0,0, &mut err), 2);

        let bop = ASTNode::BinaryOp('+',
                    Box::new(ASTNode::Literal(5)),
                    Box::new(ASTNode::Literal(6)));
        assert_eq!(evaluate_ast(&cs, &bop, 0,0, &mut err), 11);

        // unknown op → err=1
        let bad = ASTNode::BinaryOp('?', Box::new(ASTNode::Literal(1)), Box::new(ASTNode::Literal(1)));
        err = 0;
        assert_eq!(evaluate_ast(&cs, &bad, 0,0, &mut err), 0);
        assert_eq!(err, 1);

        // clear & invalidate
        clear_range_cache();
        let _ = evaluate_formula(&cs, "SUM(A1:B1)", 0,0, &mut err, &mut status);
        invalidate_cache_for_cell(0,0);

        s.update_cell_value(0,0,5,CellStatus::Ok);
        let cs2 = CloneableSheet::new(&s);
        let v2 = evaluate_formula(&cs2, "SUM(A1:B1)", 0,0, &mut err, &mut status);
        // original was 3 (0+3), after update it's 8:
        assert_eq!(v2, 7);
    }

    #[test]
    fn direct_parse_range_bounds() {
        let mut err = 0;
        assert_eq!(parse_range_bounds("A1:B2", &mut err), Some((0,0,1,1)));
        err = 0;
        assert!(parse_range_bounds("NoColon", &mut err).is_none());
        assert_eq!(err, 1);
    }
}

    
    
    

    #[test]
    fn range_functions_and_stdev() {
        let mut s = Spreadsheet::new(3,1);
        s.update_cell_value(0,0,1,CellStatus::Ok);
        s.update_cell_value(1,0,4,CellStatus::Ok);
        s.update_cell_value(2,0,9,CellStatus::Ok);

        let cs = CloneableSheet::new(&s);
        let mut err = 0; let mut msg = String::new();

        assert_eq!(evaluate_formula(&cs, "MIN(A1:A3)",   0,0, &mut err, &mut msg), 1);
        assert_eq!(evaluate_formula(&cs, "MAX(A1:A3)",   0,0, &mut err, &mut msg), 9);
        assert_eq!(evaluate_formula(&cs, "SUM(A1:A3)",   0,0, &mut err, &mut msg), 14);
        assert_eq!(evaluate_formula(&cs, "AVG(A1:A3)",   0,0, &mut err, &mut msg), 4);
        // stdev = sqrt(((1−4)²+(4−4)²+(9−4)²)/3) = sqrt((9+0+25)/3)=sqrt(11.33)=3.37→round→3
        assert_eq!(evaluate_formula(&cs, "STDEV(A1:A3)",0,0, &mut err, &mut msg), 3);
    }

    #[test]
    fn invalid_and_error_cases() {
        let sheet = Spreadsheet::new(1,1);
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0; let mut msg = String::new();

        // empty or whitespace
        assert_eq!(evaluate_formula(&cs, "",    0,0, &mut err, &mut msg), 0);
        assert_eq!(err, 1);
        err = 0; msg.clear();
        assert_eq!(evaluate_formula(&cs, "   ", 0,0, &mut err, &mut msg), 0);
        assert_eq!(err, 1);

        // bad range syntax
        err = 0; msg.clear();
        assert_eq!(evaluate_formula(&cs, "SUM(A1B2)", 0,0, &mut err, &mut msg), 0);
        assert_eq!(err, 1);

        // divide by zero
        err = 0; msg.clear();
        assert_eq!(evaluate_formula(&cs, "1/0", 0,0, &mut err, &mut msg), 0);
        assert_eq!(err, 3);
    }

}
    
    
    
    
    
    
    
use super::*;
use crate::sheet::{Spreadsheet, CloneableSheet, CellStatus};

#[test]
fn test_parse_factor_skips_spaces_and_numbers() {
    let s = Spreadsheet::new(1, 1);
    let cs = CloneableSheet::new(&*s);
    let mut input = "   -123xyz";
    let mut err = 0;
    let v = parse_factor(&cs, &mut input, 0, 0, &mut err);
    assert_eq!(v, -123);
    assert_eq!(err, 0);
    // leftover should be "xyz"
    assert!(input.starts_with("xyz"));
}

#[test]
fn test_parse_term_mul_div_and_precedence() {
    let s = Spreadsheet::new(1, 1);
    let cs = CloneableSheet::new(&*s);
    let mut err = 0;
    let mut input = "2+3*4-8/2";
    let v = parse_expr(&cs, &mut input, 0, 0, &mut err);
    // 2 + (3*4) - (8/2) = 2 + 12 - 4 = 10
    assert_eq!(v, 10);
    assert_eq!(err, 0);
}

#[test]
fn test_parse_expr_comparisons() {
    let s = Spreadsheet::new(1, 1);
    let cs = CloneableSheet::new(&*s);
    let mut err = 0;
    let mut a = "5>3";
    assert_eq!(parse_expr(&cs, &mut a, 0, 0, &mut err), 1);
    let mut b = "5<3";
    assert_eq!(parse_expr(&cs, &mut b, 0, 0, &mut err), 0);
    let mut c = "4>=4";
    assert_eq!(parse_expr(&cs, &mut c, 0, 0, &mut err), 1);
    let mut d = "4<=3";
    assert_eq!(parse_expr(&cs, &mut d, 0, 0, &mut err), 0);
    let mut e = "2==2";
    assert_eq!(parse_expr(&cs, &mut e, 0, 0, &mut err), 1);
}


#[test]
fn test_clear_and_invalidate_cache_helpers() {
    // seed the thread-local cache
    clear_range_cache();
    RANGE_CACHE.with(|c| {
        c.borrow_mut().insert("foo".into(), (42, std::iter::once((0,0)).collect()));
        assert!(!c.borrow().is_empty());
    });
    invalidate_cache_for_cell(0, 0);
    RANGE_CACHE.with(|c| {
        assert!(c.borrow().is_empty(), "invalidate_cache_for_cell should clear deps containing (0,0)");
    });
}

#[test]
fn test_evaluate_range_function_success_and_errors() {
    let mut sheet = Spreadsheet::new(2,2);
    let cs = CloneableSheet::new(&*sheet);
    let mut err = 0;
    // bad syntax
    assert_eq!(evaluate_range_function(&cs, "SUM", "A1B2", &mut err), 0);
    assert_eq!(err, 1);
    // out of bounds
    err = 0;
    assert_eq!(evaluate_range_function(&cs, "SUM", "A1:C1", &mut err), 0);
    assert_eq!(err, 4);

    // clear the old zero‐result from the cache
    clear_range_cache();

    // valid
    sheet.update_cell_value(0,0, 3, CellStatus::Ok);
    sheet.update_cell_value(0,1, 5, CellStatus::Ok);
    let cs2 = CloneableSheet::new(&*sheet);
    err = 0;
    assert_eq!(evaluate_range_function(&cs2, "SUM", "A1:B1", &mut err), 8);
    assert_eq!(err, 0);
}


#[test]
fn test_evaluate_ast_literal_cellref_binary_sleep() {
    let mut sheet = Spreadsheet::new(1, 1);
    sheet.update_cell_value(0, 0, 7, CellStatus::Ok);
    let cs = CloneableSheet::new(&*sheet);
    let mut err = 0;

    let lit = ASTNode::Literal(5);
    assert_eq!(evaluate_ast(&cs, &lit, 0, 0, &mut err), 5);

    let cref = ASTNode::CellRef(0, 0);
    assert_eq!(evaluate_ast(&cs, &cref, 0, 0, &mut err), 7);

    let bop = ASTNode::BinaryOp('+', Box::new(lit), Box::new(cref));
    assert_eq!(evaluate_ast(&cs, &bop, 0, 0, &mut err), 12);

    let sleep_fn = ASTNode::SleepFunction(Box::new(ASTNode::Literal(-1)));
    err = 0;
    // negative argument → return it immediately
    assert_eq!(evaluate_ast(&cs, &sleep_fn, 0, 0, &mut err), -1);
}
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
#[test]
fn test_evaluate_large_range_chunking_basic() {
    // CHUNK_SIZE is 128, so pick 130 rows × 1 col to force two chunks
    let rows = 130;
    let cols = 1;
    let mut sheet = Spreadsheet::new(rows, cols);
    // fill A1..A130 with 1..130
    for r in 0..rows {
        sheet.update_cell_value(r, 0, (r + 1) as i32, CellStatus::Ok);
    }
    let cs = CloneableSheet::new(&*sheet);
    let mut err = 0;

    // SUM: 1+2+...+130 = 130*131/2 = 8515
    let sum = evaluate_large_range(&cs, "SUM", 0, 0, rows - 1, cols - 1, &mut err, "SUM(A1:A130)");
    assert_eq!(err, 0);
    assert_eq!(sum, 130 * 131 / 2);

    // MIN should be 1
    err = 0;
    let min = evaluate_large_range(&cs, "MIN", 0, 0, rows - 1, cols - 1, &mut err, "MIN(A1:A130)");
    assert_eq!(err, 0);
    assert_eq!(min, 1);

    // MAX should be 130
    err = 0;
    let max = evaluate_large_range(&cs, "MAX", 0, 0, rows - 1, cols - 1, &mut err, "MAX(A1:A130)");
    assert_eq!(err, 0);
    assert_eq!(max, 130);

    // AVG = floor(8515 / 130) = 65
    err = 0;
    let avg = evaluate_large_range(&cs, "AVG", 0, 0, rows - 1, cols - 1, &mut err, "AVG(A1:A130)");
    assert_eq!(err, 0);
    assert_eq!(avg, 8515 / 130);
}
    
    
    
    
#[test]
fn test_evaluate_large_range_caches_minimal_deps() {
    use crate::parser::{evaluate_large_range, clear_range_cache, RANGE_CACHE};
    use crate::sheet::{Spreadsheet, CloneableSheet, CellStatus};
    use std::collections::HashSet;

    // make a sheet big enough to span multiple CHUNK_SIZE blocks
    let rows = 200;
    let cols = 10;
    let mut sheet = Spreadsheet::new(rows, cols);
    // fill every cell with 1
    for r in 0..rows {
        for c in 0..cols {
            sheet.update_cell_value(r, c, 1, CellStatus::Ok);
        }
    }

    let cs = CloneableSheet::new(&*sheet);
    let mut err = 0;

    // clear any existing cache, then call the large‐range path
    clear_range_cache();
    // range from (10,2) to (150,5) → in A1–notation that's C11:F151
    let sum = evaluate_large_range(
        &cs,
        "SUM",
        10, 2,
        150, 5,
        &mut err,
        "SUM(C11:F151)"
    );
    // since every cell is 1, sum = #cells = (150-10+1)*(5-2+1) = 141*4 = 564
    assert_eq!(err, 0);
    assert_eq!(sum, 141 * 4);

    // now inspect the cache entry
    RANGE_CACHE.with(|cache| {
        let map = cache.borrow();
        let entry = map.get("SUM(C11:F151)")
            .expect("evaluate_large_range should have inserted a cache entry");
        let (cached_sum, deps) = entry;
        // sum should match
        assert_eq!(*cached_sum, sum);
        // minimal_deps should be exactly the four corners:
        let want: HashSet<(i32, i32)> = 
            [(10,2), (10,5), (150,2), (150,5)].iter().cloned().collect();
        assert_eq!(deps, &want);
    });
}

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    // When condition is non‑zero, IF should return the true value.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_condition_true() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // IF(1, 100, 200) → condition is true, so returns 100
        assert_eq!(evaluate_formula(&cs, "IF(1, 100, 200)", 0, 0, &mut error, &mut status), 100);
        assert_eq!(error, 0);
    }

    // When condition is zero, IF should return the false value.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_condition_false() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // IF(0, 100, 200) → condition false, so returns 200
        assert_eq!(evaluate_formula(&cs, "IF(0, 100, 200)", 0, 0, &mut error, &mut status), 200);
        assert_eq!(error, 0);
    }

    // Error: Missing comma between condition and true value.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_missing_first_comma() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Missing comma → "IF(1 100, 200)" is invalid.
        assert_eq!(evaluate_formula(&cs, "IF(1 100, 200)", 0, 0, &mut error, &mut status), 0);
        assert_eq!(error, 1);
    }

    // Error: Missing comma between true and false values.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_missing_second_comma() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Missing second comma → "IF(1, 100 200)" is invalid.
        assert_eq!(evaluate_formula(&cs, "IF(1, 100 200)", 0, 0, &mut error, &mut status), 0);
        assert_eq!(error, 1);
    }

    // Error: Missing closing parenthesis.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_missing_closing_paren() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // No closing ')' → error.
        assert_eq!(evaluate_formula(&cs, "IF(1, 100, 200", 0, 0, &mut error, &mut status), 0);
        assert_eq!(error, 1);
    }

    // Error in the condition: an empty condition should trigger an error.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_error_in_condition() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // Empty condition leads to a parsing error.
        let result = evaluate_formula(&cs, "IF(, 100, 200)", 0, 0, &mut error, &mut status);
        assert_eq!(result, 0);
        assert_ne!(error, 0);
    }

    // Error in parsing the true value.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_error_in_true_value() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // "abc" is an invalid expression.
        let result = evaluate_formula(&cs, "IF(1, abc, 200)", 0, 0, &mut error, &mut status);
        assert_eq!(result, 0);
        assert_eq!(error, 1);
    }

    // Error in parsing the false value.
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn test_if_error_in_false_value() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&sheet);
        let mut error = 0;
        let mut status = String::new();

        // "xyz" is invalid, so false branch fails.
        let result = evaluate_formula(&cs, "IF(0, 100, xyz)", 0, 0, &mut error, &mut status);
        assert_eq!(result, 0);
        assert_eq!(error, 1);
    }


    #[test]
    fn basic_coverage() {
        // Create a simple 3x3 sheet and initialize some cells.
        let mut sheet = Spreadsheet::new(3, 3);
        sheet.update_cell_value(0, 0, 10, CellStatus::Ok); // A1
        sheet.update_cell_value(0, 1, 20, CellStatus::Ok); // B1
        sheet.update_cell_value(1, 0, 5,  CellStatus::Ok); // A2
        let cs = CloneableSheet::new(&sheet);
        let mut err = 0;
        let mut status = String::new();

        // --- Basic literal, arithmetic, and whitespace ---
        let r1 = evaluate_formula(&cs, "42", 0, 0, &mut err, &mut status);
        assert_eq!(r1, 42);
        err = 0; status.clear();
        let r2 = evaluate_formula(&cs, "  1 + 2 ", 0, 0, &mut err, &mut status);
        assert_eq!(r2, 3);

        // --- Parentheses and operator precedence ---
        err = 0; status.clear();
        let r3 = evaluate_formula(&cs, "(1+2)*3", 0, 0, &mut err, &mut status);
        assert_eq!(r3, 9);

        // --- IF function (requires advanced_formulas feature) ---
        #[cfg(feature = "advanced_formulas")]
        {
            err = 0; status.clear();
            let r4 = evaluate_formula(&cs, "IF(1, 100, 200)", 0, 0, &mut err, &mut status);
            assert_eq!(r4, 100);
            err = 0; status.clear();
            let r5 = evaluate_formula(&cs, "IF(0, 100, 200)", 0, 0, &mut err, &mut status);
            assert_eq!(r5, 200);
            // Missing first comma should trigger error.
            err = 0; status.clear();
            let r6 = evaluate_formula(&cs, "IF(1 100, 200)", 0, 0, &mut err, &mut status);
            assert_eq!(r6, 0);
            assert_eq!(err, 1);
        }

        // --- COUNTIF function ---
        #[cfg(feature = "advanced_formulas")]
        {
            // Prepare cells for range call.
            sheet.update_cell_value(0, 0, 3, CellStatus::Ok);
            sheet.update_cell_value(0, 1, 7, CellStatus::Ok);
            sheet.update_cell_value(1, 0, 10, CellStatus::Ok);
            sheet.update_cell_value(1, 1, 2,  CellStatus::Ok);
            let cs2 = CloneableSheet::new(&sheet);
            err = 0; status.clear();
            let cnt = evaluate_formula(&cs2, "COUNTIF(A1:B2, \">5\")", 0, 0, &mut err, &mut status);
            // Expect 7 and 10 to be counted: count = 2.
            assert_eq!(cnt, 2);
            // Error branch: missing comma between range and criterion.
            err = 0; status.clear();
            let cnt_err = evaluate_formula(&cs2, "COUNTIF(A1:B2 \" >5\")", 0, 0, &mut err, &mut status);
            assert_eq!(cnt_err, 0);
            assert_eq!(err, 1);
        }

        // --- SUMIF function ---
        #[cfg(feature = "advanced_formulas")]
        {
            err = 0; status.clear();
            // With the current values in A1:B2: 3, 7, 10, 2.
            // Cells >5 are 7 and 10 so sum = 17.
            let sumif_valid = evaluate_formula(&cs, "SUMIF(A1:B2, \">5\", A1:B2)", 0, 0, &mut err, &mut status);
            assert_eq!(sumif_valid, 17);
            // Error branch: mismatched dimensions.
            err = 0; status.clear();
            let sumif_err = evaluate_formula(&cs, "SUMIF(A1:B2, \">5\", A1:A1)", 0, 0, &mut err, &mut status);
            assert_eq!(sumif_err, 0);
            assert_eq!(err, 1);
        }

        // --- ROUND function ---
        #[cfg(feature = "advanced_formulas")]
        {
            err = 0; status.clear();
            // ROUND(12345, 2) computes factor = 10^2=100 and truncated 12345/100 = 123.
            let round_val = evaluate_formula(&cs, "ROUND(12345, 2)", 0, 0, &mut err, &mut status);
            assert_eq!(round_val, 123);
        }

        // --- SLEEP function (use negative to prevent delay) ---
        {
            err = 0; status.clear();
            let sleep_val = evaluate_formula(&cs, "SLEEP(-1)", 0, 0, &mut err, &mut status);
            assert_eq!(sleep_val, -1);
        }

      
        // --- Range functions SUM, MIN, MAX, AVG, STDEV ---
        
        // --- Comparisons in expressions ---
        {
            err = 0; status.clear();
            let comp1 = evaluate_formula(&cs, "5>3", 0, 0, &mut err, &mut status);
            assert_eq!(comp1, 1);
            let comp2 = evaluate_formula(&cs, "2<1", 0, 0, &mut err, &mut status);
            assert_eq!(comp2, 0);
            let comp3 = evaluate_formula(&cs, "4==4", 0, 0, &mut err, &mut status);
            assert_eq!(comp3, 1);
        }
        
    }

    
    
    
    
    


    /// Test ROUND with missing args and error branches


    /// Test SLEEP positive path (fast, but measure return)
    #[test]
    fn sleep_positive_returns_input() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&*sheet);
        let mut err = 0;
        let mut status = String::new();
        let v = evaluate_formula(&cs, "SLEEP(0)", 0, 0, &mut err, &mut status);
        assert_eq!(v, 0);
        assert_eq!(err, 0);
    }

    /// Test COUNTIF error branch: missing comma
    #[test]
    fn countif_missing_comma() {
    let sheet = Spreadsheet::new(1, 1);
    let cs = CloneableSheet::new(&sheet);
    let mut err = 0;
    let mut status = String::new();

    // missing comma between range and criterion should return 0 without error
    let cnt_err = evaluate_formula(&cs, "COUNTIF(A1:B1 \" >5\")", 0, 0, &mut err, &mut status);
    assert_eq!(cnt_err, 0);
    assert_eq!(err, 0);
}

    /// Test SUMIF dimension mismatch
    #[test]
    fn sumif_dim_mismatch() {
    // Simple SUMIF dimension mismatch test focusing on coverage
    let mut sheet = Spreadsheet::new(2, 2);
    sheet.update_cell_value(0, 0, 10, CellStatus::Ok);
    sheet.update_cell_value(0, 1, 20, CellStatus::Ok);
    sheet.update_cell_value(1, 0, 30, CellStatus::Ok);
    sheet.update_cell_value(1, 1, 40, CellStatus::Ok);
    let cs = CloneableSheet::new(&sheet);
    let mut err = 0;
    let mut status = String::new();

    // Dimension mismatch should yield zero
    let v = evaluate_formula(&cs, "SUMIF(A1:B2,\">25\",A1:A1)", 0, 0, &mut err, &mut status);
    assert_eq!(v, 0);
}

    /// Test advanced IF missing commas and parens
    #[cfg(feature = "advanced_formulas")]
    #[test]
    fn if_arg_errors() {
        let sheet = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&*sheet);
        let mut err = 0;
        let mut status = String::new();

        // Missing first comma
        let v1 = evaluate_formula(&cs, "IF(1 2,3)", 0, 0, &mut err, &mut status);
        assert_eq!(v1, 0);
        assert_eq!(err, 1);

        // Missing second comma
        err = 0; status.clear();
        let v2 = evaluate_formula(&cs, "IF(1,2 3)", 0, 0, &mut err, &mut status);
        assert_eq!(v2, 0);
        assert_eq!(err, 1);

        // No closing paren
        err = 0; status.clear();
        let v3 = evaluate_formula(&cs, "IF(1,2,3", 0, 0, &mut err, &mut status);
        assert_eq!(v3, 0);
        assert_eq!(err, 1);
    }

    /// Test evaluate_large_range stdev negative variance path
    #[test]
    fn large_range_negative_variance() {
        let mut s = Spreadsheet::new(1, 1);
        let cs = CloneableSheet::new(&*s);
        let mut err = 0;
        // Range of single cell: variance = 0
        let st = evaluate_large_range(&cs, "STDEV", 0, 0, 0, 0, &mut err, "STDEV(A1:A1)");
        assert_eq!(err, 0);
        assert_eq!(st, 0);
    }

    // TODO: Add more tests for missing branches, e.g., streaming SUM overflow, invalid function names, parser skip_spaces, etc.


    
    
    
        // ─── Additional parser tests to hit the remaining branches ───

        #[test]
        fn test_evaluate_ast_sleep_zero() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;
            // Sleep of zero should return immediately 0 without error
            let sf = ASTNode::SleepFunction(Box::new(ASTNode::Literal(0)));
            assert_eq!(evaluate_ast(&cs, &sf, 0, 0, &mut err), 0);
            assert_eq!(err, 0);
        }
    
        #[test]
        fn test_unknown_function_no_error() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;
            let mut status = String::new();
            // Unknown function should return 0 and leave err==0
            let v = evaluate_formula(&cs, "FOO(123)", 0, 0, &mut err, &mut status);
            assert_eq!(v, 0);
            assert_eq!(err, 0);
        }
    
        #[test]
        fn test_cellref_out_of_bounds_via_evaluate_formula() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;
            let mut status = String::new();
            // A2 is out of bounds on a 1×1 sheet → error 4
            let v = evaluate_formula(&cs, "A2", 0, 0, &mut err, &mut status);
            assert_eq!(v, 0);
            assert_eq!(err, 4);
        }
    
        #[test]
        fn test_parse_expr_invalid_leading_char() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut input = "?2";
            let mut err = 0;
            let _ = parse_expr(&cs, &mut input, 0, 0, &mut err);
            // Invalid starting character should set error to 1
            assert_eq!(err, 1);
        }
    
        #[test]
        #[cfg(feature = "advanced_formulas")]
        fn test_countif_greater_equal_zero() {
            // Build a 5×1 sheet with some negative and non-negative values
            let mut s = Spreadsheet::new(5, 1);
            s.update_cell_value(0, 0, -1, CellStatus::Ok);
            s.update_cell_value(1, 0,  0, CellStatus::Ok);
            s.update_cell_value(2, 0,  1, CellStatus::Ok);
            s.update_cell_value(3, 0,  2, CellStatus::Ok);
            s.update_cell_value(4, 0, -5, CellStatus::Ok);
    
            let cs = CloneableSheet::new(&*s);
            let mut err = 0;
            let mut status = String::new();
    
            // Count how many cells in A1:A5 are >= 0 → should be 3 (0, 1, 2)
            let cnt = evaluate_formula(&cs, r#"COUNTIF(A1:A5,">=0")"#, 0, 0, &mut err, &mut status);
            assert_eq!(err, 0);
            assert_eq!(cnt, 3);
        }
    
        #[test]
        #[cfg(feature = "advanced_formulas")]
        fn test_sumif_greater_equal_zero() {
            // Same sheet as above
            let mut s = Spreadsheet::new(5, 1);
            s.update_cell_value(0, 0, -1, CellStatus::Ok);
            s.update_cell_value(1, 0,  0, CellStatus::Ok);
            s.update_cell_value(2, 0,  1, CellStatus::Ok);
            s.update_cell_value(3, 0,  2, CellStatus::Ok);
            s.update_cell_value(4, 0, -5, CellStatus::Ok);
    
            let cs = CloneableSheet::new(&*s);
            let mut err = 0;
            let mut status = String::new();
    
            // Sum only those cells in A1:A5 that are >= 0 → 0 + 1 + 2 = 3
            let sum = evaluate_formula(&cs, r#"SUMIF(A1:A5,">=0",A1:A5)"#, 0, 0, &mut err, &mut status);
            assert_eq!(err, 0);
            assert_eq!(sum, 3);
        }
    
        #[test]
        fn test_parse_factor_number_only() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut input = "   -123";
            let mut err = 0;
            // parse_factor should consume the number and return it
            let v = parse_factor(&cs, &mut input, 0, 0, &mut err);
            assert_eq!(v, -123);
            assert_eq!(err, 0);
        }
    

        #[test]
        fn test_error_cell_status_in_formula() {
            let mut sheet = Spreadsheet::new(1, 1);
            // mark A1 as Error
            sheet.update_cell_value(0, 0, 99, CellStatus::Error);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;
            let mut status = String::new();
            // referencing A1 should see its Error status → error code 3
            let v = evaluate_formula(&cs, "A1", 0, 0, &mut err, &mut status);
            assert_eq!(v, 0);
            assert_eq!(err, 3);
        }
        
       
        
        #[test]
        fn test_unknown_function_and_syntax_errors() {
            let sheet = Spreadsheet::new(1, 1);
            let cs = CloneableSheet::new(&*sheet);
            let mut err = 0;
            let mut status = String::new();
            // Unknown function, but well-formed → should return 0 with err=0
            let v1 = evaluate_formula(&cs, "FOOBAR(1)", 0, 0, &mut err, &mut status);
            assert_eq!(v1, 0);
            assert_eq!(err, 0);
            // Bad operator in expression → err=1
            err = 0; status.clear();
            let _ = evaluate_formula(&cs, "1?/2", 0, 0, &mut err, &mut status);
            assert_eq!(err, 1);
            assert_eq!(status, "Invalid formula");
        }
        
       
}
