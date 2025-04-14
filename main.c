#include <stdio.h>
#include <stdlib.h>
#include <ctype.h>
#include <string.h>
#include <sys/time.h>
#include "sheet.h"
#include "parser.h"

#define VIEWPORT_SIZE 10

bool check= false;
/* --- Scrolling Functions --- */
void col_to_letters(int col, char *buffer) {
    int len = 0;
    do {
        buffer[len++] = 'A' + (col % 26);
        col = col / 26 - 1;
    } while (col >= 0);
    buffer[len] = '\0';
    for (int i = 0; i < len / 2; i++) {
        char tmp = buffer[i];
        buffer[i] = buffer[len - 1 - i];
        buffer[len - 1 - i] = tmp;
    }
}

int cell_to_indices(const char *cell, int *row, int *col) {
    return cellNameToCoords(cell, row, col);
}

void clamp_viewport_ve(int totalRows, int *start_row){
    if (*start_row>totalRows) 
    (*start_row)-=VIEWPORT_SIZE;
    else if ((*start_row)>(totalRows-VIEWPORT_SIZE))
    *start_row=totalRows-VIEWPORT_SIZE;
    else if (*start_row<0) *start_row=0;

}

void clamp_viewport_hz(int totalCols,int *start_col) {
    // Horizontal bounds
    if (*start_col>totalCols) 
    (*start_col)-=VIEWPORT_SIZE;
    else if ((*start_col)>(totalCols-VIEWPORT_SIZE))
    *start_col=totalCols-VIEWPORT_SIZE;
    else if (*start_col<0) *start_col=0;
    
}

void display_grid(const Spreadsheet *sheet) {
    int startRow = sheet->topRow;
    int startCol = sheet->leftCol;
    int endRow = startRow + VIEWPORT_SIZE;
    int endCol = startCol + VIEWPORT_SIZE;

    if (endRow > sheet->totalRows) endRow = sheet->totalRows;
    if (endCol > sheet->totalCols) endCol = sheet->totalCols;

    // Print column headers
    printf("     "); // Space for row numbers
    for (int c = startCol; c < endCol; c++) {
        char col_buf[16];
        col_to_letters(c, col_buf);
        printf("%-12s", col_buf); // Ensure column header is aligned
    }
    printf("\n");

    // Print rows with values
    for (int r = startRow; r < endRow; r++) {
        printf("%-4d ", r + 1); // Row number
        for (int c = startCol; c < endCol; c++) {
            if (sheet->cells[r][c].status == CELL_ERROR) {
                printf("%-12s", "ERR");
            } else {
                printf("%-12d", sheet->cells[r][c].value);
            }
        }
        printf("\n");
    }
}
void display_grid_from(const Spreadsheet *sheet, int startRow, int startCol) {
    // Print column headers
    printf("     ");
    int max_col = startCol + VIEWPORT_SIZE;
    if (max_col > sheet->totalCols) max_col = sheet->totalCols; //Prevent printing headers beyond the grid

    for(int c = startCol; c < max_col; c++) {
        
        char col_buf[16];
        col_to_letters(c, col_buf);
        printf("%-12s", col_buf);
    }
    printf("\n");

    // Print rows with data
    int max_row = startRow + VIEWPORT_SIZE;
        if (max_row > sheet->totalRows) max_row = sheet->totalRows; //Prevent printing rows beyond the grid

    for(int r = startRow; r < max_row; r++) {
        // Row number (1-based)
        printf("%-4d ", r + 1);

        // Cell values
        for(int c = startCol; c < max_col; c++) {
            if (sheet->cells[r][c].status == CELL_ERROR)
                printf("%-12s", "ERR");
            else
                printf("%-12d", sheet->cells[r][c].value);
        }
        printf("\n");
    }

}



/* --- Command Processing --- */
void process_command(Spreadsheet *sheet, const char *cmd , char *statusMsg) {

    if (strcmp(cmd, "w") == 0) {
        sheet->topRow -= VIEWPORT_SIZE;
        clamp_viewport_ve(sheet->totalRows, &sheet->topRow);
    } else if (strcmp(cmd, "s") == 0) {
        sheet->topRow += VIEWPORT_SIZE;
        clamp_viewport_ve(sheet->totalRows, &sheet->topRow);
    } else if (strcmp(cmd, "a") == 0) {
        sheet->leftCol -= VIEWPORT_SIZE;
        clamp_viewport_hz(sheet->totalCols, &sheet->leftCol);
    } else if (strcmp(cmd, "d") == 0) {
        sheet->leftCol += VIEWPORT_SIZE;
        clamp_viewport_hz(sheet->totalCols, &sheet->leftCol);
    } else if (strncmp(cmd, "scroll_to", 9) == 0) {
        char cellName[20];
        if (sscanf(cmd, "scroll_to %s", cellName) == 1) {
            int row, col;
            if (!cellNameToCoords(cellName, &row, &col)) {
                strcpy(statusMsg, "Invalid cell");
            }
            else if(row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
                strcpy(statusMsg, "Cell reference out of bounds");
            }
             else {
                sheet->topRow = row;
                sheet->leftCol = col;
            }
        } else {
            strcpy(statusMsg, "Invalid command");
        }
    } else if (strcmp(cmd, "disable_output") == 0) {
        sheet->outputEnabled = 0;
    } else if (strcmp(cmd, "enable_output") == 0) {
        sheet->outputEnabled = 1;
    } else if (strchr(cmd, '=') != NULL) { // Cell assignment
        char cellName[20];
        const char *expr = strchr(cmd, '=') + 1;
        strncpy(cellName, cmd, strchr(cmd, '=') - cmd);
        cellName[strchr(cmd, '=') - cmd] = '\0';
        
        int row, col;
        if (!cellNameToCoords(cellName, &row, &col)) {
            strcpy(statusMsg, "Invalid cell");
        } else if (row < 0 || row >= sheet->totalRows || col < 0 || col >= sheet->totalCols) {
            strcpy(statusMsg, "Cell out of bounds");
        } else {
            updateCellFormula(sheet, row, col, expr, statusMsg);
        }
    } else {
        strcpy(statusMsg, "unrecognized cmd");
    }

}


int main(int argc, char *argv[]) {
    if (argc != 3) {
        fprintf(stderr, "Usage: %s <rows> <cols>\n", argv[0]);
        return EXIT_FAILURE;
    }
    int rows = atoi(argv[1]);
    int cols = atoi(argv[2]);
    if (rows < 1 || cols < 1) {
        fprintf(stderr, "Invalid dimensions.\n");
        return EXIT_FAILURE;
    }
    
    char cmd[256];
    char statusMsg[30] = "ok";
    struct timeval start, end;
    double elapsedTime = 0.0;

    //Time to make the spreadsheet
    // gettimeofday(&start, NULL);
    Spreadsheet *sheet = createSpreadsheet(rows, cols);
    // gettimeofday(&end, NULL);

    // elapsedTime = (end.tv_sec - start.tv_sec) + (end.tv_usec - start.tv_usec) / 1000000.0;

    display_grid(sheet);
    printf("[%.1f] (%s) > ", elapsedTime, statusMsg);

    while (1) {
        if (fgets(cmd, sizeof(cmd), stdin) == NULL) {
            strcpy(statusMsg, "Invalid command");
        }
        size_t len = strlen(cmd);
        if (len > 0 && cmd[len - 1] == '\n')
            cmd[len - 1] = '\0';
        
        if (strcmp(cmd, "q") == 0)
            break;
        
        //Process the command and calculate time

        gettimeofday(&start, NULL);
        process_command(sheet, cmd,statusMsg);
        gettimeofday(&end, NULL);

        elapsedTime = (end.tv_sec - start.tv_sec) + (end.tv_usec - start.tv_usec) / 1000000.0;
        if (sheet->outputEnabled && !check)
            display_grid_from(sheet,sheet->topRow,sheet->leftCol);
        printf("[%.1f] (%s) > ", elapsedTime, statusMsg);
        strcpy(statusMsg, "ok");
    }
    
    freeSpreadsheet(sheet);
    return EXIT_SUCCESS;
}
