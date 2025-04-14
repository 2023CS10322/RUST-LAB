#ifndef SHEET_H
#define SHEET_H

#include <stdlib.h>
#include <stdbool.h>

/* Display dimensions for scrolling */
#define DISP_ROWS 10
#define DISP_COLS 10

/* Enumeration for cell status */
typedef enum {
    CELL_OK,     // Valid cell value.
    CELL_ERROR   // Error (e.g., division by zero)
} CellStatus;

/* Forward declaration of Spreadsheet and Cell */
struct Spreadsheet;
struct Cell;

/* Structure representing one cell in the spreadsheet */
typedef struct Cell {
    int value;          // Computed integer value
    char *formula;      // Formula string (if any)
    CellStatus status;  // CELL_OK or CELL_ERROR

    /* Dependency tracking using direct pointers to other cells */
    struct Cell **dependencies;   // Array of pointers to cells that this cell depends on
    int numDependencies;          // Number of dependencies

    struct Cell **dependents;     // Array of pointers to cells that depend on this cell
    int numDependents;            // Number of dependents

    /* Coordinates of this cell (for convenience) */
    int row;
    int col;

    /* Pointer back to the parent spreadsheet (useful for dependency lookups) */
    struct Spreadsheet *sheet;
} Cell;

/* Main spreadsheet structure */
typedef struct Spreadsheet {
    int totalRows;      // Total number of rows
    int totalCols;      // Total number of columns
    Cell **cells;       // 2D array of pointers to cells
    int topRow;         // Top row index for current view
    int leftCol;        // Left column index for current view
    int outputEnabled;  // 1 = display output; 0 = suppressed

    /* Temporary buffer for selective recalculation (if desired) */
    int *tempIndegree;
    int tempArraySize;
    int skipDefaultDisplay;
} Spreadsheet;

/* Function prototypes */
Spreadsheet* createSpreadsheet(int rows, int cols);
void freeSpreadsheet(Spreadsheet *sheet);
void displaySpreadsheet(const Spreadsheet *sheet);
void updateCellFormula(Spreadsheet *sheet, int row, int col, const char *formula, char *statusMsg);
bool recalcAll(Spreadsheet *sheet,char *statusMsg);  // Legacy full recalculation (for reference)

/* Dependency helper functions */
void clearDependencies(Cell *cell);
void addDependency(Cell *cell, int depRow, int depCol);
void addDependencyPointer(Cell *cell, Cell *depCell);
void addDependent(Cell *cell, int depRow, int depCol);
void addDependentPointer(Cell *cell, Cell *depCell);
void removeDependent(Cell *cell, int depRow, int depCol);

/* Utility functions for cell name conversion */
int cellNameToCoords(const char *name, int *row, int *col);
void coordsToCellName(int row, int col, char *buffer, size_t bufSize);

#endif
