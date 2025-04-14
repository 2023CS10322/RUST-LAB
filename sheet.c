#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "sheet.h"
#include "parser.h"
#include <stdbool.h>
#include <math.h>

/* 
   Updated extractDependencies:
   Now returns an array of pointers to Cells referenced in the formula.
*/
/* Helper: trim leading and trailing whitespace from a string */
int cellNameToCoords(const char *name, int *row, int *col)
{
    int pos = 0, colVal = 0;
    while (name[pos] && isalpha(name[pos])) {
        colVal = colVal * 26 + (toupper(name[pos]) - 'A' + 1);
        pos++;
    }
    if (colVal == 0)
        return 0;
    *col = colVal - 1;
    
    int rowVal = 0;
    while (name[pos] && isdigit(name[pos])) {
        rowVal = rowVal * 10 + (name[pos] - '0');
        pos++;
    }
    if (rowVal <= 0)
        return 0;
    *row = rowVal - 1;
    
    if (name[pos] != '\0')
        return 0;
    return 1;
    
}
void trim(char *str) {
    char *end;
    // trim leading space
    while(isspace((unsigned char)*str)) str++;
    if(*str == 0) {  // All spaces?
        str[0] = '\0';
        return;
    }
    // trim trailing space
    end = str + strlen(str) - 1;
    while(end > str && isspace((unsigned char)*end)) end--;
    // Write new null terminator
    *(end+1) = '\0';
}

/* The validformula function.
   Returns 0 if the formula is valid; returns 1 and sets statusMsg otherwise.
*/
int validformula(Spreadsheet* sheet, const char* formula, char* statusMsg) {
    /* Clear statusMsg */
    statusMsg[0] = '\0';

    int len = strlen(formula);
    if(len == 0) {
        strcpy(statusMsg, "Empty formula");
        return 1;
    }
    /* Check if the entire formula is a cell reference.
       If so, validate that it is in bounds.
    */
    {
        int row, col;
        if(cellNameToCoords(formula, &row, &col)) {
            if(row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
                strcpy(statusMsg, "Cell reference out of bounds");
                return 1;
            }
            return 0;  // Valid formula as it is a correct cell reference.
        }
    }

    
    {
        char *endptr;
        strtol(formula, &endptr, 10);
        if (*endptr == '\0') {
            // The whole string was parsed as an integer.
            return 0;
        }
    }

    /* Check for function formulas: Format 5 (range functions) or Format 6 (SLEEP) */
    if (strncmp(formula, "MAX(", 4) == 0 ||
        strncmp(formula, "MIN(", 4) == 0 ||
        strncmp(formula, "SUM(", 4) == 0 ||
        strncmp(formula, "AVG(", 4) == 0 ||
        strncmp(formula, "STDEV(", 6) == 0) {
        // Identify the function name by finding the '('
        int pos = 0;
        while(formula[pos] != '(' && formula[pos] != '\0' && pos < 15) {
            pos++;
        }
        if(formula[pos] != '(') {
            strcpy(statusMsg, "Missing '(' after function name");
            return 1;
        }
        // Check that the formula ends with ')'
        if(formula[len-1] != ')') {
            strcpy(statusMsg, "Missing closing parenthesis");
            return 1;
        }
        /* For range functions (MAX, MIN, SUM, AVG, STDEV), we expect:
             FUNC(cellRef1:cellRef2)
           where cellRef1 and cellRef2 are separated by a colon.
        */
        char inner[128] = {0};
        strncpy(inner, formula + pos + 1, len - pos - 2);
        inner[len - pos - 2] = '\0';
        trim(inner);
        char *colon = strchr(inner, ':');
        if (!colon) {
            strcpy(statusMsg, "Missing colon in range");
            return 1;
        }
        *colon = '\0';
        char cell1[64] = {0}, cell2[64] = {0};
        strncpy(cell1, inner, sizeof(cell1)-1);
        strncpy(cell2, colon+1, sizeof(cell2)-1);
        trim(cell1);
        trim(cell2);

        int row1, col1, row2, col2;
        if (!(cellNameToCoords(cell1, &row1, &col1)) != 0) {
            strcpy(statusMsg, "Invalid first cell reference");
            return 1;
        }
        if (!(cellNameToCoords(cell2, &row2, &col2)) != 0) {
            strcpy(statusMsg, "Invalid second cell reference");
            return 1;
        }
        if (row1 < 0 || row1 >= sheet->totalRows || col1 < 0 || col1 >= sheet->totalCols) {
            strcpy(statusMsg, "First cell reference out of bounds");
            return 1;
        }
        if (row2 < 0 || row2 >= sheet->totalRows || col2 < 0 || col2 >= sheet->totalCols) {
            strcpy(statusMsg, "Second cell reference out of bounds");
            return 1;
        }
        /* Check that the first cell is top-left of the second */
        if (row1 > row2 || col1 > col2) {
            strcpy(statusMsg, "Invalid range order");
            return 1;
        }
        return 0;
    } else if (strncmp(formula, "SLEEP(", 6) == 0) {
        /* Format 6: SLEEP(x) where x is either an integer or a cell reference.
           Check that the formula ends with ')'.
        */
        if(formula[len-1] != ')') {
            strcpy(statusMsg, "Missing closing parenthesis in SLEEP");
            return 1;
        }
        char inner[128] = {0};
        strncpy(inner, formula + 6, len - 7);
        inner[len-7] = '\0';
        trim(inner);
        int sleepInt;
        if (sscanf(inner, "%d", &sleepInt) == 1) {
            // It is an integer (which can be positive or negative)
            return 0;
        } else {
            // Try to interpret inner as a cell reference.
            int row, col;
            if (!(cellNameToCoords(inner, &row, &col)) != 0) {
                strcpy(statusMsg, "Invalid cell reference in SLEEP");
                return 1;
            }
            if (row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
                strcpy(statusMsg, "Cell reference in SLEEP out of bounds");
                return 1;
            }
            return 0;
        }
    }
    
    /* Otherwise, we expect one of the binary operations.
       The valid binary formats are:
         1. integer op cellReference    e.g. 4+A4 or -1+A4
         2. cellReference op integer    e.g. A4+4
         3. integer op integer          e.g. 4+5 or -1+2
         4. cellReference op cellReference   e.g. A4+B5
       The operator (op) must be one of: +, -, *, /
    */
    int opIndex = -1;
    int i = (formula[0]=='-' ? 1 : 0); // Skip initial '-' if present.
    for (; formula[i] != '\0'; i++) {
        if (formula[i] == '+' || formula[i] == '-' ||
            formula[i] == '*' || formula[i] == '/') {
            opIndex = i;
            break;
        }
    }
    if (opIndex == -1) {
        strcpy(statusMsg, "Operator not found");
        return 1;
    }
    
    char left[128] = {0}, right[128] = {0};
    strncpy(left, formula, opIndex);
    left[opIndex] = '\0';
    strcpy(right, formula + opIndex + 1);
    trim(left);
    trim(right);
    
    int leftInt, rightInt;
    int isLeftInt = (sscanf(left, "%d", &leftInt) == 1);
    int isRightInt = (sscanf(right, "%d", &rightInt) == 1);
    int leftIsCell = 0, rightIsCell = 0;
    int row, col;
    
    if (!isLeftInt) {
        if (!(cellNameToCoords(left, &row, &col)) == 0) {
            if (row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
                strcpy(statusMsg, "Left cell reference out of bounds");
                return 1;
            }
            leftIsCell = 1;
        }
    }
    
    if (!isRightInt) {
        if (!(cellNameToCoords(right, &row, &col)) == 0) {
            if (row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
                strcpy(statusMsg, "Right cell reference out of bounds");
                return 1;
            }
            rightIsCell = 1;
        }
    }
    
    /* The valid cases are when each side is either an integer or a valid cell reference */
    if ((isLeftInt || leftIsCell) && (isRightInt || rightIsCell))
        return 0;
    
    strcpy(statusMsg, "Invalid formula format");
    return 1;
}
void extractDependencies(Spreadsheet *sheet, const char *formula, Cell ***deps, int *numDeps) {
    *deps = NULL;
    *numDeps = 0;
    int capacity = 0;
    const char *p = formula;
    while (*p) {
        // Skip any non-alphabetic characters
        while (*p && !isalpha((unsigned char)*p))
            p++;
        if (!*p)
            break;
        const char *start = p;
        // Read letters (column)
        while (*p && isalpha((unsigned char)*p))
            p++;
        // Read digits (row)
        while (*p && isdigit((unsigned char)*p))
            p++;
        // Check for range (colon)
        if (*p == ':') {
            p++; // skip colon
            const char *rangeStart2 = p;
            while (*p && isalpha((unsigned char)*p))
                p++;
            while (*p && isdigit((unsigned char)*p))
                p++;
            int len1 = (int)(strchr(start, ':') - start);
            if (len1 >= 20)
                len1 = 19;
            char cellRef1[20];
            strncpy(cellRef1, start, len1);
            cellRef1[len1] = '\0';

            int len2 = (int)(p - rangeStart2);
            if (len2 >= 20)
                len2 = 19;
            char cellRef2[20];
            strncpy(cellRef2, rangeStart2, len2);
            cellRef2[len2] = '\0';

            int startRow, startCol, endRow, endCol;
            if (!cellNameToCoords(cellRef1, &startRow, &startCol) ||
                !cellNameToCoords(cellRef2, &endRow, &endCol))
                continue;
            // Ensure forward order
            if (startRow > endRow) {
                int tmp = startRow;
                startRow = endRow;
                endRow = tmp;
            }
            if (startCol > endCol) {
                int tmp = startCol;
                startCol = endCol;
                endCol = tmp;
            }
            for (int rr = startRow; rr <= endRow; rr++) {
                for (int cc = startCol; cc <= endCol; cc++) {
                    if (*numDeps >= capacity) {
                        capacity = (capacity == 0) ? 4 : capacity * 2;
                        *deps = realloc(*deps, capacity * sizeof(Cell*));
                        if (!*deps) {
                            fprintf(stderr, "Memory allocation error in extractDependencies.\n");
                            exit(EXIT_FAILURE);
                        }
                    }
                    (*deps)[*numDeps] = &(sheet->cells[rr][cc]);
                    (*numDeps)++;
                }
            }
        }
        else {
            int len = (int)(p - start);
            if (len >= 20)
                len = 19;
            char cellRef[20];
            strncpy(cellRef, start, len);
            cellRef[len] = '\0';
            int r, c;
            if (!cellNameToCoords(cellRef, &r, &c))
                continue;
            if (*numDeps >= capacity) {
                capacity = (capacity == 0) ? 4 : capacity * 2;
                *deps = realloc(*deps, capacity * sizeof(Cell*));
                if (!*deps) {
                    fprintf(stderr, "Memory allocation error in extractDependencies.\n");
                    exit(EXIT_FAILURE);
                }
            }
            (*deps)[*numDeps] = &(sheet->cells[r][c]);
            (*numDeps)++;
        }
    }
}

/* -------------------------------
   Dependency Graph Management (using direct Cell pointers)
   ------------------------------- */

// Clears the dependencies array for a cell.
void clearDependencies(Cell *cell) {
    if (cell->dependencies) {
        free(cell->dependencies);
        cell->dependencies = NULL;
    }
    cell->numDependencies = 0;
}

// Adds a dependency to a cell given coordinates.
void addDependency(Cell *cell, int depRow, int depCol) {
    Cell *depCell = &(cell->sheet->cells[depRow][depCol]);
    addDependencyPointer(cell, depCell);
}

// Adds a dependency using a direct pointer.
void addDependencyPointer(Cell *cell, Cell *depCell) {
    cell->dependencies = realloc(cell->dependencies, (cell->numDependencies + 1) * sizeof(Cell *));
    if (!cell->dependencies) {
        fprintf(stderr, "Memory allocation error in addDependencyPointer.\n");
        exit(EXIT_FAILURE);
    }
    cell->dependencies[cell->numDependencies] = depCell;
    cell->numDependencies++;
}

// Adds a dependent to a cell given coordinates.
void addDependent(Cell *cell, int depRow, int depCol) {
    Cell *depCell = &(cell->sheet->cells[depRow][depCol]);
    addDependentPointer(cell, depCell);
}

// Adds a dependent using a direct pointer.
void addDependentPointer(Cell *cell, Cell *depCell) {
    cell->dependents = realloc(cell->dependents, (cell->numDependents + 1) * sizeof(Cell *));
    if (!cell->dependents) {
        fprintf(stderr, "Memory allocation error in addDependentPointer.\n");
        exit(EXIT_FAILURE);
    }
    cell->dependents[cell->numDependents] = depCell;
    cell->numDependents++;
}

// Removes a dependent from a cell’s dependents array (by comparing row and col).
void removeDependent(Cell *cell, int depRow, int depCol) {
    for (int i = 0; i < cell->numDependents; i++) {
        Cell *d = cell->dependents[i];
        if (d->row == depRow && d->col == depCol) {
            cell->dependents[i] = cell->dependents[cell->numDependents - 1];
            cell->numDependents--;
            return;
        }
    }
}

/* -------------------------------
   Circular Dependency Detection
   ------------------------------- */

/*
   hasCycle: Recursively checks if starting from 'cell' following dependency edges
   eventually reaches 'target'. The visited matrix (indexed by cell coordinates) avoids re‐processing.
*/
static bool hasCycle(Cell *cell, Cell *target, bool **visited) {
    // For each dependency of the current cell
    for (int i = 0; i < cell->numDependencies; i++) {
        Cell *dep = cell->dependencies[i];
        if (dep == target)
            return true;
        if (!visited[dep->row][dep->col]) {
            visited[dep->row][dep->col] = true;
            if (hasCycle(dep, target, visited))
                return true;
        }
    }
    return false;
}

/*
   hasCircularDependency:
   Returns true if the given cell (following its dependency pointers) eventually depends on itself.
*/
static bool hasCircularDependency(Cell *cell) {
    int rows = cell->sheet->totalRows;
    int cols = cell->sheet->totalCols;
    bool **visited = malloc(rows * sizeof(bool *));
    if (!visited) {
        fprintf(stderr, "Memory allocation error in hasCircularDependency.\n");
        exit(EXIT_FAILURE);
    }
    for (int i = 0; i < rows; i++) {
        visited[i] = calloc(cols, sizeof(bool));
        if (!visited[i]) {
            fprintf(stderr, "Memory allocation error in hasCircularDependency (row).\n");
            exit(EXIT_FAILURE);
        }
    }
    bool result = hasCycle(cell, cell, visited);
    for (int i = 0; i < rows; i++)
        free(visited[i]);
    free(visited);
    return result;
}

/* -------------------------------
   Selective Recalculation Using DFS
   ------------------------------- */

/*
   dfsCollect: Recursively traverse the dependency graph (using dependents)
   and collect all affected cells (those that depend on the updated cell).
   The starting cell (start) is not added.
*/
// static void dfsCollect(Spreadsheet *sheet, Cell *cell,
//                        bool **visited, Cell ***affected, int *count, int *capacity, Cell *start) {
//     int r = cell->row, c = cell->col;
//     if (visited[r][c])
//         return;
//     visited[r][c] = true;
//     if (cell != start) {
//         if (*count >= *capacity) {
//             *capacity = (*capacity == 0) ? 10 : (*capacity * 2);
//             *affected = realloc(*affected, (*capacity) * sizeof(Cell *));
//             if (!*affected) {
//                 fprintf(stderr, "Memory allocation error in dfsCollect.\n");
//                 exit(EXIT_FAILURE);
//             }
//         }
//         (*affected)[*count] = cell;
//         (*count)++;
//     }
//     for (int i = 0; i < cell->numDependents; i++) {
//         dfsCollect(sheet, cell->dependents[i], visited, affected, count, capacity, start);
//     }
// }

static void dfsCollect(Spreadsheet *sheet, Cell *start, 
                       bool **visited, Cell ***affected, int *count, int *capacity) {
    // Stack for iterative DFS
    int stackSize = 100; // Initial stack size
    Cell **stack = malloc(stackSize * sizeof(Cell *));
    if (!stack) {
        fprintf(stderr, "Memory allocation error in dfsCollect.\n");
        exit(EXIT_FAILURE);
    }
    int top = -1;

    // Push the start cell
    stack[++top] = start;

    while (top >= 0) {
        // Pop a cell from the stack
        Cell *cell = stack[top--];

        int r = cell->row, c = cell->col;
        if (visited[r][c]) 
            continue;
        
        visited[r][c] = true;

        if (cell != start) {
            // Resize affected array if necessary
            if (*count >= *capacity) {
                *capacity = (*capacity == 0) ? 10 : (*capacity * 2);
                *affected = realloc(*affected, (*capacity) * sizeof(Cell *));
                if (!*affected) {
                    fprintf(stderr, "Memory allocation error in dfsCollect.\n");
                    exit(EXIT_FAILURE);
                }
            }
            (*affected)[*count] = cell;
            (*count)++;
        }

        // Push dependents onto the stack
        for (int i = 0; i < cell->numDependents; i++) {
            if (top + 1 >= stackSize) {
                stackSize *= 2;
                stack = realloc(stack, stackSize * sizeof(Cell *));
                if (!stack) {
                    fprintf(stderr, "Memory allocation error in dfsCollect.\n");
                    exit(EXIT_FAILURE);
                }
            }
            stack[++top] = cell->dependents[i];
        }
    }

    free(stack);
}
/*
   recalcAffected: Recalculate only those cells (in the affected set) that depend
   (directly or indirectly) on the updated cell at (startRow, startCol).
   This function collects the affected cells via DFS, builds an indegree array for just
   those cells, and then processes them in topological order.
*/
void recalcAffected(Spreadsheet *sheet, int startRow, int startCol, char * statusMsg) {
    Cell *start = &(sheet->cells[startRow][startCol]);
    
    // Allocate and initialize visited matrix.
    bool **visited = malloc(sheet->totalRows * sizeof(bool *));
    if (!visited) {
        fprintf(stderr, "Memory allocation error in recalcAffected (visited).\n");
        exit(EXIT_FAILURE);
    }
    for (int i = 0; i < sheet->totalRows; i++) {
        visited[i] = calloc(sheet->totalCols, sizeof(bool));
        if (!visited[i]) {
            fprintf(stderr, "Memory allocation error in recalcAffected (visited row).\n");
            exit(EXIT_FAILURE);
        }
    }
    
    // Collect affected cells via DFS.
    Cell **affected = NULL;
    int count = 0, capacity = 0;
    // dfsCollect(sheet, start, visited, &affected, &count, &capacity, start);
    dfsCollect(sheet, start, visited, &affected, &count, &capacity);

    // Free visited matrix.
    for (int i = 0; i < sheet->totalRows; i++)
        free(visited[i]);
    free(visited);
    
    if (count == 0) {
        free(affected);
        return;
    }
    
    // Allocate an indegree array for the affected set.
    int *indegree = malloc(count * sizeof(int));
    if (!indegree) {
        fprintf(stderr, "Memory allocation error in recalcAffected (indegree).\n");
        exit(EXIT_FAILURE);
    }
    for (int i = 0; i < count; i++)
        indegree[i] = 0;
    
    // For each affected cell, count dependencies that are also in the affected set.
    for (int i = 0; i < count; i++) {
        Cell *cell = affected[i];
        for (int j = 0; j < cell->numDependencies; j++) {
            Cell *dep = cell->dependencies[j];
            for (int k = 0; k < count; k++) {
                if (affected[k] == dep) {
                    indegree[i]++;
                    break;
                }
            }
        }
    }
    
    // Allocate a queue for topological sorting.
    int *queue = malloc(count * sizeof(int));
    if (!queue) {
        fprintf(stderr, "Memory allocation error in recalcAffected (queue).\n");
        exit(EXIT_FAILURE);
    }
    int front = 0, rear = 0;
    for (int i = 0; i < count; i++) {
        if (indegree[i] == 0)
            queue[rear++] = i;
    }
    
    // Process the affected cells in topological order.
    while (front < rear) {
        int idx = queue[front++];
        Cell *cell = affected[idx];
        if (cell->formula != NULL) {
            int errorFlag = 0;
            int newVal = evaluateFormula(sheet, cell->formula, cell->row, cell->col, &errorFlag,statusMsg);
            if (errorFlag==3)
                cell->status = CELL_ERROR;
            else if (errorFlag == 2) {
                // This is your newly distinguished "invalid range" error
                strcpy(statusMsg, "Invalid range");
                return;
            } 
            else if (errorFlag == 1) {
                // Some other error (e.g. bad parse, etc.)
                strcpy(statusMsg, "Error in formula");
                return;
            } 
            else {
                cell->value = newVal;
                cell->status = CELL_OK;
            }
        }
        for (int i = 0; i < cell->numDependents; i++) {
            Cell *dep = cell->dependents[i];
            for (int k = 0; k < count; k++) {
                if (affected[k] == dep) {
                    indegree[k]--;
                    if (indegree[k] == 0)
                        queue[rear++] = k;
                    break;
                }
            }
        }
    }
    
    free(indegree);
    free(queue);
    free(affected);
}

/* -------------------------------
   Legacy Full Recalculation (for reference)
   ------------------------------- */
// bool recalcAll(Spreadsheet *sheet, char * statusMsg) {
//     int totalCells = sheet->totalRows * sheet->totalCols;
//     int *indegree = calloc(totalCells, sizeof(int));
//     if (!indegree) {
//         fprintf(stderr, "Memory allocation error in recalcAll.\n");
//         exit(EXIT_FAILURE);
//     }
//     for (int r = 0; r < sheet->totalRows; r++) {
//         for (int c = 0; c < sheet->totalCols; c++) {
//             int idx = r * sheet->totalCols + c;
//             if (sheet->cells[r][c].formula != NULL) {
//                 int cnt = 0;
//                 for (int k = 0; k < sheet->cells[r][c].numDependencies; k++) {
//                     Cell *dep = sheet->cells[r][c].dependencies[k];
//                     if (dep->formula != NULL)
//                         cnt++;
//                 }
//                 indegree[idx] = cnt;
//             }
//         }
//     }
//     int *queue = malloc(totalCells * sizeof(int));
//     if (!queue) {
//         fprintf(stderr, "Memory allocation error in recalcAll.\n");
//         exit(EXIT_FAILURE);
//     }
//     int front = 0, rear = 0;
//     for (int r = 0; r < sheet->totalRows; r++) {
//         for (int c = 0; c < sheet->totalCols; c++) {
//             int idx = r * sheet->totalCols + c;
//             if (sheet->cells[r][c].formula != NULL && indegree[idx] == 0)
//                 queue[rear++] = idx;
//         }
//     }
//     int countOrder = 0;
//     int *order = malloc(totalCells * sizeof(int));
//     if (!order) {
//         fprintf(stderr, "Memory allocation error in recalcAll.\n");
//         exit(EXIT_FAILURE);
//     }
//     while (front < rear) {
//         int idx = queue[front++];
//         order[countOrder++] = idx;
//         int r = idx / sheet->totalCols;
//         int c = idx % sheet->totalCols;
//         Cell *cell = &sheet->cells[r][c];
//         for (int i = 0; i < cell->numDependents; i++) {
//             int depIdx = cell->dependents[i]->row * sheet->totalCols + cell->dependents[i]->col;
//             if (sheet->cells[cell->dependents[i]->row][cell->dependents[i]->col].formula != NULL) {
//                 indegree[depIdx]--;
//                 if (indegree[depIdx] == 0)
//                     queue[rear++] = depIdx;
//             }
//         }
//     }
//     int formulaCount = 0;
//     for (int r = 0; r < sheet->totalRows; r++) {
//         for (int c = 0; c < sheet->totalCols; c++) {
//             if (sheet->cells[r][c].formula != NULL)
//                 formulaCount++;
//         }
//     }
//     bool cycleDetected = (countOrder != formulaCount);
//     if (cycleDetected) {
//         free(indegree);
//         free(queue);
//         free(order);
//         return false;
//     }
//     for (int i = 0; i < countOrder; i++) {
//         int idx = order[i];
//         int r = idx / sheet->totalCols;
//         int c = idx % sheet->totalCols;
//         if (sheet->cells[r][c].formula != NULL) {
//             int errorFlag = 0;
//             int newVal = evaluateFormula(sheet, sheet->cells[r][c].formula, r, c, &errorFlag,statusMsg);
//             sheet->cells[r][c].value = newVal;
//             sheet->cells[r][c].status = CELL_OK;
//         }
//     }
//     free(indegree);
//     free(queue);
//     free(order);
//     return true;
// }

/* -------------------------------
   Spreadsheet Creation and Freeing
   ------------------------------- */
Spreadsheet* createSpreadsheet(int rows, int cols) {
    Spreadsheet *sheet = malloc(sizeof(Spreadsheet));
    if (!sheet) {
        fprintf(stderr, "Error: Memory allocation failed.\n");
        exit(EXIT_FAILURE);
    }
    sheet->totalRows = rows;
    sheet->totalCols = cols;
    sheet->topRow = 0;
    sheet->leftCol = 0;
    sheet->outputEnabled = 1;
    sheet->tempIndegree = NULL;
    sheet->tempArraySize = 0;
    sheet->skipDefaultDisplay = 0;
    
    sheet->cells = malloc(rows * sizeof(Cell *));
    if (!sheet->cells) {
        fprintf(stderr, "Error: Memory allocation failed.\n");
        exit(EXIT_FAILURE);
    }
    for (int r = 0; r < rows; r++) {
        sheet->cells[r] = malloc(cols * sizeof(Cell));
        if (!sheet->cells[r]) {
            fprintf(stderr, "Error: Memory allocation failed.\n");
            exit(EXIT_FAILURE);
        }
        for (int c = 0; c < cols; c++) {
            Cell *cell = &sheet->cells[r][c];
            cell->value = 0;
            cell->formula = NULL;
            cell->status = CELL_OK;
            cell->dependencies = NULL;
            cell->numDependencies = 0;
            cell->dependents = NULL;
            cell->numDependents = 0;
            cell->row = r;
            cell->col = c;
            cell->sheet = sheet;
        }
    }
    return sheet;
}

void freeSpreadsheet(Spreadsheet *sheet) {
    if (!sheet)
        return;
    for (int r = 0; r < sheet->totalRows; r++) {
        for (int c = 0; c < sheet->totalCols; c++) {
            if (sheet->cells[r][c].formula)
                free(sheet->cells[r][c].formula);
            if (sheet->cells[r][c].dependencies)
                free(sheet->cells[r][c].dependencies);
            if (sheet->cells[r][c].dependents)
                free(sheet->cells[r][c].dependents);
        }
        free(sheet->cells[r]);
    }
    free(sheet->cells);
    if (sheet->tempIndegree)
        free(sheet->tempIndegree);
    free(sheet);
}

void markCellAndDependentsAsError(Cell *cell) {
    if (cell->status == CELL_ERROR)
        return; // Already marked as an error

    cell->status = CELL_ERROR;
    cell->value = 0;

    for (int i = 0; i < cell->numDependents; i++) {
        markCellAndDependentsAsError(cell->dependents[i]);
    }
}

/* -------------------------------
   Updating a Cell's Formula and Recalculating Dependencies
   ------------------------------- */
void updateCellFormula(Spreadsheet *sheet, int row, int col, const char *formula , char * statusMsg) {
    int volidi;
    volidi=validformula(sheet,formula,statusMsg);
    if (volidi==1)
    {
        strcpy(statusMsg,"Unrecognized");
        return;
    }
    strcpy(statusMsg,"Ok");
    Cell *cell = &sheet->cells[row][col];
    // BACKUP: Save the current dependencies.
    int oldNumDeps = cell->numDependencies;
    Cell **oldDeps = NULL;
    if (oldNumDeps > 0) {
        oldDeps = malloc(oldNumDeps * sizeof(Cell *));
        if (!oldDeps) {
            fprintf(stderr, "Memory allocation error in updateCellFormula (oldDeps backup).\n");
            exit(EXIT_FAILURE);
        }
        for (int i = 0; i < oldNumDeps; i++) {
            oldDeps[i] = cell->dependencies[i];
        }
    }
    // BACKUP: Save the old formula.
    char *oldFormula = cell->formula ? strdup(cell->formula) : NULL;

    // 1) Remove old dependencies: remove this cell from each dependency's dependents list.
    for (int i = 0; i < cell->numDependencies; i++) {
        Cell *dep = cell->dependencies[i];
        removeDependent(dep, cell->row, cell->col);
    }
    clearDependencies(cell);

    // 2) Store new formula string.
    if (cell->formula)
        free(cell->formula);
    cell->formula = strdup(formula);

    // 3) Check if the formula is a constant.
    bool isConstant = true;
    for (int i = 0; formula[i]; i++) {
        if (!isdigit((unsigned char)formula[i]) && formula[i] != '-')
            isConstant = false;
    }

    if (!isConstant) {
        // Parse new dependencies.
        Cell **newDeps = NULL;
        int numNewDeps = 0;
        extractDependencies(sheet, formula, &newDeps, &numNewDeps);
        // Add them.
        for (int i = 0; i < numNewDeps; i++) {

            addDependencyPointer(cell, newDeps[i]);
            addDependentPointer(newDeps[i], cell);
        }
        if (newDeps)
            free(newDeps);
    }

    // 4) Detect circular dependency.
    if (hasCircularDependency(cell)) {
        char cellName[20];
        coordsToCellName(row, col, cellName, sizeof(cellName));
        strcpy(statusMsg, "Circular dependency detected in cell ");
        strcat(statusMsg, cellName);

        // Remove any newly added dependencies.
        for (int i = 0; i < cell->numDependencies; i++) {
            Cell *dep = cell->dependencies[i];
            removeDependent(dep, cell->row, cell->col);
        }
        clearDependencies(cell);

        // Restore the old formula.
        if (oldFormula) {
            if (cell->formula)
                free(cell->formula);
            cell->formula = oldFormula;
        } else {
            cell->formula = NULL;
        }

        // Restore the old dependencies.
        for (int i = 0; i < oldNumDeps; i++) {
            addDependencyPointer(cell, oldDeps[i]);
            addDependentPointer(oldDeps[i], cell);
        }

        free(oldDeps);
        return;
    }
    //B1/C1
    free(oldDeps);
    if (oldFormula)
        free(oldFormula);
    // 5) Evaluate the formula.
    int errorFlag = 0;
    int newVal = evaluateFormula(sheet, cell->formula, row, col, &errorFlag,statusMsg);
    if (errorFlag == 3) { // Division by zero
        markCellAndDependentsAsError(cell);
        strcpy(statusMsg, "Ok");
        return;
    }
    else if (errorFlag==4)
    {
        strcpy(statusMsg, "Range out of bounds");
        return;
    }
    else {
        // No error
        cell->value = newVal;
        cell->status = CELL_OK;
        // 6) Recalculate only affected cells if no error occurred.
        
        recalcAffected(sheet, row, col,statusMsg);
    }
    return;
}

/* -------------------------------
   Utility Functions
   ------------------------------- */


void coordsToCellName(int row, int col, char *buffer, size_t bufSize) {
    char colStr[10];
    int n = col + 1;
    int pos = 0;
    while (n > 0) {
        int rem = (n - 1) % 26;
        colStr[pos++] = 'A' + rem;
        n = (n - 1) / 26;
    }
    colStr[pos] = '\0';
    for (int i = 0; i < pos / 2; i++) {
        char tmp = colStr[i];
        colStr[i] = colStr[pos - 1 - i];
        colStr[pos - 1 - i] = tmp;
    }
    snprintf(buffer, bufSize, "%s%d", colStr, row + 1);
}
