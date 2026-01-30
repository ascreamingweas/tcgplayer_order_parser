# TCGplayer Packing Slip Organizer
# Makefile for common operations

# Input PDF path (can be anywhere on disk)
PDF ?=
# Output filename (optional, auto-generated from PDF name if not specified)
OUTPUT ?=

# Output directory for generated HTML files
OUTPUT_DIR := output

# Python settings
VENV := venv
PYTHON := $(VENV)/bin/python3
PIP := $(VENV)/bin/pip

.PHONY: help setup run clean list

help: ## Show this help message
	@echo "TCGplayer Packing Slip Organizer"
	@echo ""
	@echo "Usage:"
	@echo "  make setup                    - Create virtual environment and install dependencies"
	@echo "  make run PDF=/path/to/file.pdf - Process a packing slip PDF (full path supported)"
	@echo "  make list                     - List generated HTML files in output folder"
	@echo "  make clean                    - Remove all generated HTML files"
	@echo ""
	@echo "Examples:"
	@echo "  make run PDF=~/Downloads/TCGplayer_PackingSlips.pdf"
	@echo "  make run PDF=/Users/me/Documents/orders/packing_slip.pdf"
	@echo "  make run PDF=~/Downloads/slip.pdf OUTPUT=my_orders.html"
	@echo ""
	@echo "Output files are saved to: ./$(OUTPUT_DIR)/"

setup: ## Create virtual environment and install dependencies
	@if [ ! -d "$(VENV)" ]; then \
		echo "Creating virtual environment..."; \
		python3 -m venv $(VENV); \
	fi
	@echo "Installing dependencies..."
	@$(PIP) install --quiet pdfplumber
	@mkdir -p $(OUTPUT_DIR)
	@echo "Setup complete!"

$(OUTPUT_DIR):
	@mkdir -p $(OUTPUT_DIR)

run: $(OUTPUT_DIR) ## Process a packing slip PDF (requires PDF=<path>)
	@if [ -z "$(PDF)" ]; then \
		echo "Error: Please specify a PDF file"; \
		echo "Usage: make run PDF=<path/to/file.pdf>"; \
		exit 1; \
	fi
	@EXPANDED_PDF=$$(eval echo "$(PDF)"); \
	if [ ! -f "$$EXPANDED_PDF" ]; then \
		echo "Error: File '$$EXPANDED_PDF' not found"; \
		exit 1; \
	fi; \
	if [ ! -d "$(VENV)" ]; then \
		echo "Virtual environment not found. Run 'make setup' first."; \
		exit 1; \
	fi; \
	if [ -z "$(OUTPUT)" ]; then \
		BASENAME=$$(basename "$$EXPANDED_PDF" .pdf); \
		OUTPUT_FILE="$(OUTPUT_DIR)/$${BASENAME}_organized.html"; \
	else \
		OUTPUT_FILE="$(OUTPUT_DIR)/$(OUTPUT)"; \
	fi; \
	$(PYTHON) mtg_packing_slip_organizer.py "$$EXPANDED_PDF" "$$OUTPUT_FILE"

list: ## List generated HTML files in output folder
	@echo "Generated HTML files in $(OUTPUT_DIR)/:"
	@if [ -d "$(OUTPUT_DIR)" ]; then \
		ls -1 $(OUTPUT_DIR)/*.html 2>/dev/null || echo "  No HTML files found"; \
	else \
		echo "  Output directory does not exist yet"; \
	fi

clean: ## Remove all generated HTML files
	@echo "Removing generated HTML files from $(OUTPUT_DIR)/..."
	@rm -rf $(OUTPUT_DIR)/*.html 2>/dev/null || true
	@echo "Done."
