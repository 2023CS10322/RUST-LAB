#include <stdio.h>
#include <stdlib.h>
#include <ctype.h>
#include <string.h>
#include <unistd.h>
#include <math.h>
#include "parser.h"

// Forward declarations for recursive descent parsing
static int parseExpr(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error);
static int parseTerm(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error);
static int parseFactor(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error);

static void skipSpaces(const char **inputPtr) {
    while (**inputPtr && isspace((unsigned char)**inputPtr))
        (*inputPtr)++;
}

/*
   evaluateRangeFunction:
   Evaluates one of: MIN, MAX, SUM, AVG, STDEV on a range like "A1:B10".
   If the range is invalid or reversed, we set *error to 2 (invalid range).
   If any cell in the range is in error, we set *error to 3, etc.
*/
static int evaluateRangeFunction(Spreadsheet *sheet, const char *funcName, const char *rangeStr, int *error) {
    const char *colonPos = strchr(rangeStr, ':');
    if (colonPos == NULL) {
        *error = 1;  // missing colon => invalid
        return 0;
    }

    int len1 = colonPos - rangeStr;
    char cell1[20];
    if (len1 >= (int)sizeof(cell1)) { 
        *error = 1; 
        return 0; 
    }
    strncpy(cell1, rangeStr, len1);
    cell1[len1] = '\0';
    // Trim trailing spaces from cell1
    for (int i = strlen(cell1) - 1; i >= 0 && isspace((unsigned char)cell1[i]); i--)
        cell1[i] = '\0';

    char cell2[20];
    strncpy(cell2, colonPos + 1, sizeof(cell2) - 1);
    cell2[sizeof(cell2) - 1] = '\0';
    // Trim trailing spaces from cell2
    int l = (int)strlen(cell2);
    while (l > 0 && isspace((unsigned char)cell2[l - 1])) {
        cell2[l - 1] = '\0';
        l--;
    }

    int startRow, startCol, endRow, endCol;
    if (!cellNameToCoords(cell1, &startRow, &startCol) ||
        !cellNameToCoords(cell2, &endRow, &endCol)) {
        *error = 1;
        return 0;
    }

    // Check that the range is forward (top-left to bottom-right).
    if (startRow > endRow || startCol > endCol) {
        *error = 2;  // "invalid range"
        return 0;
    }

    long long sum = 0;
    int minVal = 2147483647;    // INT_MAX
    int maxVal = -2147483648;   // INT_MIN
    int count = 0;

    for (int r = startRow; r <= endRow; r++) {
        for (int c = startCol; c <= endCol; c++) {
            if (sheet->cells[r][c].status == CELL_ERROR) {
                *error = 3; // a cell in range is in error
                return 0;
            }
            int value = sheet->cells[r][c].value;
            sum += value;
            if (value < minVal) minVal = value;
            if (value > maxVal) maxVal = value;
            count++;
        }
    }
    if (count == 0) {
        *error = 1;  // empty range => invalid
        return 0;
    }

    if (strcmp(funcName, "MIN") == 0) {
        return minVal;
    } else if (strcmp(funcName, "MAX") == 0) {
        return maxVal;
    } else if (strcmp(funcName, "SUM") == 0) {
        return (int)sum;
    } else if (strcmp(funcName, "AVG") == 0) {
        return (int)(sum / count);
    } else if (strcmp(funcName, "STDEV") == 0) {
        // double mean = (double)sum / count;
        // double varianceSum = 0.0;
        // if (n <= 1) return 0;  // Avoid division by zero

        int sum = 0,count=0, mean;
        double variance = 0.0;
        for (int rr = startRow; rr <= endRow; rr++) {
            for (int cc = startCol; cc <= endCol; cc++) {
                sum+=(sheet->cells[rr][cc].value);
                count+=1;
                // double diff = sheet->cells[rr][cc].value - mean;
                // varianceSum += diff * diff;
            }
        }
        mean=sum/count;
        for (int rr = startRow; rr <= endRow; rr++) {
            for (int cc = startCol; cc <= endCol; cc++) {
                // sum+=(sheet->cells[rr][cc].value);
                // count+=1;
                // double diff = sheet->cells[rr][cc].value - mean;
                // varianceSum += diff * diff;
                variance += (sheet->cells[rr][cc].value - mean) * (sheet->cells[rr][cc].value - mean);
            }
        }
        variance /= count;
        // double variance = varianceSum / count;
        // Return integer standard deviation (rounded)
        return (int)round(sqrt(variance));
    }

    *error = 1;  // unrecognized function
    return 0;
}

static int parseExpr(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error) {
    int result = parseTerm(sheet, inputPtr, curRow, curCol, error);
    if (*error) return 0;

    skipSpaces(inputPtr);
    while (**inputPtr == '+' || **inputPtr == '-') {
        char op = **inputPtr;
        (*inputPtr)++;
        skipSpaces(inputPtr);
        int termValue = parseTerm(sheet, inputPtr, curRow, curCol, error);
        if (*error) return 0;
        if (op == '+') result += termValue;
        else           result -= termValue;
        skipSpaces(inputPtr);
    }

    // Final skip of spaces:
    skipSpaces(inputPtr);

    // Allow parseExpr to end if we hit a closing parenthesis (belongs to a higher-level call).
    // Only treat leftover text as error if it's neither the end of the string nor a valid closing parenthesis.
    if (**inputPtr != '\0' && **inputPtr != ')') {
        // If the leftover isn't just whitespace, mark error
        if (!isspace((unsigned char)**inputPtr)) {
            *error = 1;  // leftover junk => invalid
        }
    }

    return result;
}

static int parseTerm(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error) {
    int value = parseFactor(sheet, inputPtr, curRow, curCol, error);
    if (*error) return 0;

    skipSpaces(inputPtr);
    while (**inputPtr == '*' || **inputPtr == '/') {
        char op = **inputPtr;
        (*inputPtr)++;
        skipSpaces(inputPtr);

        int factorValue = parseFactor(sheet, inputPtr, curRow, curCol, error);
        if (*error) return 0;

        if (op == '/') {
            if (factorValue == 0) {
                *error = 3;  // division by zero
                return 0;
            }
            value /= factorValue;
        } else { // '*'
            value *= factorValue;
        }
        skipSpaces(inputPtr);
    }
    return value;
}

static int parseFactor(Spreadsheet *sheet, const char **inputPtr, int curRow, int curCol, int *error) {
    skipSpaces(inputPtr);

    // Case 1: function or cell reference (starts with a letter)
    if (isalpha(**inputPtr)) {
        // Read the token (e.g. MAX, SLEEP, or cell reference).
        const char *start = *inputPtr;
        while (**inputPtr && isalpha(**inputPtr))
            (*inputPtr)++;
        int tokenLength = (int)(*inputPtr - start);
        char token[20];
        if (tokenLength >= (int)sizeof(token)) {
            *error = 1;
            return 0;
        }
        strncpy(token, start, tokenLength);
        token[tokenLength] = '\0';

        skipSpaces(inputPtr);
        // If next char is '(', it's a function call
        if (**inputPtr == '(') {
            (*inputPtr)++;  // skip '('
            skipSpaces(inputPtr);

            // Check which function
            if (strcmp(token, "SLEEP") == 0) {
                int sleepTime = parseExpr(sheet, inputPtr, curRow, curCol, error);
                if (*error) return 0;
                skipSpaces(inputPtr);
                if (**inputPtr == ')') {
                    (*inputPtr)++;
                }
                // If negative => do not sleep, just return negative
                if (sleepTime < 0) {
                    return sleepTime;
                } else {
                    sleep(sleepTime);
                    return sleepTime;
                }
            }
            else if (!strcmp(token, "MIN")  || !strcmp(token, "MAX")  ||
                     !strcmp(token, "SUM")  || !strcmp(token, "AVG")  ||
                     !strcmp(token, "STDEV")) {
                // Range function
                const char *rangeStart = *inputPtr;
                const char *closeParen = strchr(rangeStart, ')');
                if (!closeParen) {
                    *error = 1; // missing ')'
                    return 0;
                }
                int rangeLen = (int)(closeParen - rangeStart);
                char *rangeStr = malloc(rangeLen + 1);
                if (!rangeStr) {
                    *error = 1;
                    return 0;
                }
                strncpy(rangeStr, rangeStart, rangeLen);
                rangeStr[rangeLen] = '\0';

                int val = evaluateRangeFunction(sheet, token, rangeStr, error);
                free(rangeStr);

                // skip the ')'
                *inputPtr = closeParen;
                if (**inputPtr == ')') {
                    (*inputPtr)++;
                }
                return val;
            }
            else {
                // Some unknown function => skip until ')'
                while (**inputPtr && **inputPtr != ')')
                    (*inputPtr)++;
                if (**inputPtr == ')') {
                    (*inputPtr)++;
                }
                // Return 0 for unknown function
                return 0;
            }
        }
        else {
            // Not a function call => treat as a cell reference
            // We already advanced past the letters, so let's revert
            *inputPtr = start;
            char cellRef[20];
            int pos = 0;
            while (**inputPtr && (isalnum((unsigned char)**inputPtr)) && pos < (int)sizeof(cellRef) - 1) {
                cellRef[pos++] = *(*inputPtr)++;
            }
            cellRef[pos] = '\0';

            int r, c;
            if (!cellNameToCoords(cellRef, &r, &c)) {
                *error = 1;
                return 0;
            }
            if (r < 0 || r >= sheet->totalRows || c < 0 || c >= sheet->totalCols) {
                *error = 4;/////major change
                return 0;
            }
            if (sheet->cells[r][c].status == CELL_ERROR) {
                *error = 3;
                return 0;
            }
            return sheet->cells[r][c].value;
        }
    }

    // Case 2: number (possibly negative)
    if (isdigit(**inputPtr) || ((**inputPtr == '-') && isdigit(*(*inputPtr + 1)))) {
        int sign = 1;
        if (**inputPtr == '-') {
            sign = -1;
            (*inputPtr)++;
        }
        int number = 0;
        while (isdigit((unsigned char)**inputPtr)) {
            number = number * 10 + (**inputPtr - '0');
            (*inputPtr)++;
        }
        return sign * number;
    }

    // Case 3: parenthesized expression
    if (**inputPtr == '(') {
        (*inputPtr)++; // skip '('
        int val = parseExpr(sheet, inputPtr, curRow, curCol, error);
        if (*error) return 0;
        if (**inputPtr == ')') {
            (*inputPtr)++;
        }
        return val;
    }

    // Otherwise => error
    *error = 1;
    return 0;
}

int evaluateFormula(Spreadsheet *sheet, const char *formula, int currentRow, int currentCol, int *error, char *statusMsg) {
    char *trimmed = strdup(formula);
    if (!trimmed) {
        *error = 1;
        strcpy(statusMsg, "Memory allocation error");
        return 0;
    }
    // Trim leading spaces
    char *start = trimmed;
    while (*start && isspace((unsigned char)*start)) {
        start++;
    }
    memmove(trimmed, start, strlen(start) + 1);

    // Trim trailing spaces
    int len = (int)strlen(trimmed);
    while (len > 0 && isspace((unsigned char)trimmed[len - 1])) {
        trimmed[len - 1] = '\0';
        len--;
    }

    const char *ptr = trimmed;
    *error = 0;
    int result = parseExpr(sheet, &ptr, currentRow, currentCol, error);

    if (*error == 1) {
        strcpy(statusMsg, "Invalid formula");
        free(trimmed);
        return 0;
    }
    else if (*error == 2) {
        // Means invalid range
        strcpy(statusMsg, "Invalid range");
        free(trimmed);
        return 0;
    }
    else if (*error == 3) {
        // e.g. division by zero or cell error
        free(trimmed);
        return 0; // We'll handle that in updateCellFormula
    }

    free(trimmed);
    return result;
}
