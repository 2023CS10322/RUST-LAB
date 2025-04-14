#ifndef PARSER_H
#define PARSER_H

#include "sheet.h"

/* 
   Evaluate a formula string in the context of the spreadsheet.
   currentRow and currentCol indicate the cell's location.
   On error (e.g., division by zero), *error is set to nonzero.
*/
int evaluateFormula(Spreadsheet *sheet, const char *formula, int currentRow, int currentCol, int *error, char *statusMsg);

#endif
