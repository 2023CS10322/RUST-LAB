#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int main(void) {
    FILE *fp = fopen("input.txt", "w");
    if (!fp) {
        perror("fopen");
        return 1;
    }
    
    // Disable output initially to avoid flooding the screen.
    fprintf(fp, "disable_output\n");
    
    // 1. Scroll to top-left for a known starting point.
    fprintf(fp, "scroll_to A1\n");
    
    // 2. Simple assignment.
    fprintf(fp, "A1=100\n");
    
    // 3. Arithmetic addition.
    fprintf(fp, "B1=A1+50\n");
    
    // 4. Arithmetic subtraction (may become negative).
    fprintf(fp, "C1=A1-B1\n");
    
    // 5. Multiplication.
    fprintf(fp, "D1=A1*B1\n");
    
    // 6. Division.
    fprintf(fp, "E1=A1/2\n");
    
    // 7. Division by zero.
    fprintf(fp, "F1=A1/0\n");
    
    // 8. Range function: MIN over A1 and B1.
    fprintf(fp, "G1=MIN(A1:B1)\n");
    
    // 9. Range function: MAX over A1 and B1.
    fprintf(fp, "H1=MAX(A1:B1)\n");
    
    // 10. Range function: SUM over A1, B1, C1.
    fprintf(fp, "I1=SUM(A1:C1)\n");
    
    // 11. Range function: AVG over A1, B1, C1.
    fprintf(fp, "J1=AVG(A1:C1)\n");
    
    // 12. Range function: STDEV over A1, B1, C1.
    fprintf(fp, "K1=STDEV(A1:C1)\n");
    
    // 13-15. Chain dependency.
    fprintf(fp, "L1=A1+1\n");  // L1 depends on A1.
    fprintf(fp, "M1=L1+1\n");  // M1 depends on L1.
    fprintf(fp, "N1=M1+L1\n"); // N1 depends on both M1 and L1.
    
    // 16-17. Circular dependency between O1 and P1.
    fprintf(fp, "O1=P1+1\n");
    fprintf(fp, "P1=O1+1\n");
    
    // 18. Out-of-bounds reference (assuming sheet dimensions: 1000 rows x 2000 cols).
    fprintf(fp, "Q1=Z1000+1\n");
    
    // 19. SLEEP with a valid argument (sleeps for 1 second).
    fprintf(fp, "R1=SLEEP(1)\n");
    
    // 20. SLEEP with a negative argument (should not sleep, returns negative).
    fprintf(fp, "S1=SLEEP(-3)\n");
    
    // 21. Advanced range function: SUM from A1 to K1.
    fprintf(fp, "T1=SUM(A1:K1)\n");
    
    // 22. Advanced range function: AVG from A1 to K1.
    fprintf(fp, "U1=AVG(A1:K1)\n");
    
    // 23. Complex arithmetic expression with parentheses.
    fprintf(fp, "V1=(A1+B1)*(C1-D1)/E1\n");
    
    // 24. Formula with extra spaces.
    fprintf(fp, "W1 =  10   +   20\n");
    
    // 25. Self-reference (should trigger circular dependency error).
    fprintf(fp, "X1=X1+1\n");
    
    // 26. Unknown function.
    fprintf(fp, "Y1=FOO(A1)\n");
    
    // 27. Arithmetic referencing an ERR cell (if any earlier error occurred).
    fprintf(fp, "Z1=E1+F1\n");
    
    // 28. Overwriting an existing cell with a new formula.
    fprintf(fp, "A1=SUM(A1:B1)\n");
    
    // 29. Range function over a single row.
    fprintf(fp, "AA1=AVG(A1:E1)\n");
    
    // 30. Negative constant only.
    fprintf(fp, "AB1=-50\n");
    
    // Finally, re-enable output so that some output is displayed.
    fprintf(fp, "enable_output\n");
    
    // Quit the program.
    fprintf(fp, "q\n");
    
    fclose(fp);
    
    // Run the spreadsheet program with the test input.
    // Adjust the path and dimensions if needed.
    int ret = system("./target/release/spreadsheet 1000 2000 < input.txt > output.txt");
    if (ret != 0) {
        printf("Error running the spreadsheet program.\n");
        return 1;
    }
    
    printf("Advanced edge-case tests complete. Check output.txt for results.\n");
    return 0;
}
