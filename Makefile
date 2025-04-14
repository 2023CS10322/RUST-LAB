# Compiler and flags
CC = gcc
CFLAGS = -Wall -Wextra -std=c99

# Source and object files for the main executable
SRCS = main.c parser.c sheet.c
OBJS = $(SRCS:.c=.o)

# Directory for the final binary
TARGET_DIR = target/release
TARGET = $(TARGET_DIR)/spreadsheet

# Test source and executable
TEST_SRCS = test.c
TEST_OBJS = $(TEST_SRCS:.c=.o)
TEST_TARGET = $(TARGET_DIR)/test

# Report file (LaTeX source)
REPORT_SRC = COP290_Lab_Report.tex
REPORT_PDF = report.pdf

.PHONY: all clean test report

# Default target: build the main executable
all: $(TARGET)

$(TARGET): $(OBJS)
	@mkdir -p $(TARGET_DIR)
	$(CC) $(CFLAGS) -o $(TARGET) $(OBJS) -lm

%.o: %.c
	$(CC) $(CFLAGS) -c $< -o $@

# Test target: build test executable and run tests
test: $(TEST_TARGET)
	./$(TEST_TARGET)

$(TEST_TARGET): $(TEST_OBJS) $(TARGET)
	@mkdir -p $(TARGET_DIR)
	$(CC) $(CFLAGS) -o $(TEST_TARGET) $(TEST_OBJS) -lm

# Report target: generate report.pdf from LaTeX source
report: $(REPORT_PDF)

$(REPORT_PDF): $(REPORT_SRC)
	pdflatex -jobname=report $(REPORT_SRC)
	pdflatex -jobname=report $(REPORT_SRC)

# Clean target: remove object files, executables, and temporary files
clean:
	rm -f $(OBJS) $(TEST_OBJS)
	rm -f $(TARGET) $(TEST_TARGET)
	rm -f $(REPORT_PDF) COP290_Lab_Report.pdf
	rm -f *.aux *.log *.toc
