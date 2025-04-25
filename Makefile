# Makefile for the Rust spreadsheet project

# Use Cargo as the build tool.
CARGO = cargo
TARGET = target/release/sheet

# Default rows/cols (can be overridden on the command line)
ROWS ?= 10
COLS ?= 10

.PHONY: all clean test report coverage ext1 ext2 docs
# Default target: build the main executable in release mode.
all: $(TARGET)

$(TARGET):
	$(CARGO) build --release

# Test target: run all tests.
test:
	$(CARGO) test

# Report target: generate report.pdf from the LaTeX source.
# Report target: open the existing PDF report in your project root.
report:
	@echo "Opening report.pdf…"
	open report.pdf
 coverage:
	$(CARGO) tarpaulin --lib --ignore-tests

# Extension 1: full‐featured TUI (history, undo, advanced formulas)
ext1:
	$(CARGO) run --features cli_full -- $(ROWS) $(COLS)

# Extension 2: GUI frontend
ext2:
	$(CARGO) run --no-default-features --features gui_app -- $(ROWS) $(COLS)
# bring up both the HTML api docs and your PDF report in one shot
docs: 
	$(CARGO) doc --no-deps --open
	$(MAKE) report



# Clean target: remove build artifacts and temporary files.
clean:
	$(CARGO) clean
	rm -f *.aux *.log *.toc
