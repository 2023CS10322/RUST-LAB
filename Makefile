# Makefile for the Rust spreadsheet project

# Use Cargo as the build tool.
CARGO = cargo
TARGET = target/release/sheet

.PHONY: all clean test report

# Default target: build the main executable in release mode.
all: $(TARGET)

$(TARGET):
	$(CARGO) build --release

# Test target: run all tests.
test:a
	$(CARGO) test

# Report target: generate report.pdf from the LaTeX source.
report:
	pdflatex -jobname=report COP290_Lab_Report.tex
	pdflatex -jobname=report COP290_Lab_Report.tex

# Clean target: remove build artifacts and temporary files.
clean:
	$(CARGO) clean
	rm -f report.pdf *.aux *.log *.toc
