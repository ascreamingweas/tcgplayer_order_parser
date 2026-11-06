# TCGplayer Packing Slip Organizer
# Makefile for common operations

# Input PDF path; PDF2/PDF3 add more for a multi-order pull sheet
PDF ?=
# Additional PDFs for multi-order pull sheets (optional)
PDF2 ?=
PDF3 ?=
# TCGplayer pull sheet CSV export (alternative to PDFs; covers all open orders)
CSV ?=
# TCGplayer shipping export CSV for the same orders (optional; adds bulk labels/packing)
SHIP ?=
# Port for the web app
PORT ?= 8000
# Output filename (optional, auto-generated from PDF name if not specified)
OUTPUT ?=

# Output directory for generated HTML files
OUTPUT_DIR := output

# Python settings
VENV := venv
PYTHON := $(VENV)/bin/python3
PIP := $(VENV)/bin/pip

.PHONY: help setup run pullsheet serve test clean list

help: ## Show this help message
	@echo "TCGplayer Packing Slip Organizer"
	@echo ""
	@echo "Usage:"
	@echo "  make setup                    - Create virtual environment and install dependencies"
	@echo "  make run PDF=/path/to/file.pdf - Process a packing slip PDF (full path supported)"
	@echo "  make pullsheet CSV=/path/to/pullsheet.csv - Process a TCGplayer pull sheet CSV"
	@echo "  make serve                    - Run the web app (pull, labels, and packing) on PORT=$(PORT)"
	@echo "  make test                     - Run the tests"
	@echo "  make list                     - List generated HTML files in output folder"
	@echo "  make clean                    - Remove all generated HTML files"
	@echo ""
	@echo "Examples:"
	@echo "  make run PDF=~/Downloads/TCGplayer_PackingSlips.pdf"
	@echo "  make run PDF=~/Downloads/slip.pdf OUTPUT=my_orders.html"
	@echo "  make run PDF=order1.pdf PDF2=order2.pdf PDF3=order3.pdf  (multi-order)"
	@echo "  make pullsheet CSV=~/Downloads/TCGplayer_PullSheet.csv"
	@echo "  make pullsheet CSV=~/Downloads/PullSheet.csv SHIP=~/Downloads/ShippingExport.csv"
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
	EXTRA_PDFS=""; \
	if [ -n "$(PDF2)" ]; then \
		EXPANDED_PDF2=$$(eval echo "$(PDF2)"); \
		EXTRA_PDFS="$$EXTRA_PDFS \"$$EXPANDED_PDF2\""; \
	fi; \
	if [ -n "$(PDF3)" ]; then \
		EXPANDED_PDF3=$$(eval echo "$(PDF3)"); \
		EXTRA_PDFS="$$EXTRA_PDFS \"$$EXPANDED_PDF3\""; \
	fi; \
	$(PYTHON) mtg_packing_slip_organizer.py "$$EXPANDED_PDF" $$EXTRA_PDFS -o "$$OUTPUT_FILE"

pullsheet: $(OUTPUT_DIR) ## Process a TCGplayer pull sheet CSV (requires CSV=<path>)
	@if [ -z "$(CSV)" ]; then \
		echo "Error: Please specify a CSV file"; \
		echo "Usage: make pullsheet CSV=<path/to/pullsheet.csv>"; \
		exit 1; \
	fi
	@EXPANDED_CSV=$$(eval echo "$(CSV)"); \
	if [ ! -f "$$EXPANDED_CSV" ]; then \
		echo "Error: File '$$EXPANDED_CSV' not found"; \
		exit 1; \
	fi; \
	if [ ! -d "$(VENV)" ]; then \
		echo "Virtual environment not found. Run 'make setup' first."; \
		exit 1; \
	fi; \
	if [ -z "$(OUTPUT)" ]; then \
		BASENAME=$$(basename "$$EXPANDED_CSV" .csv); \
		OUTPUT_FILE="$(OUTPUT_DIR)/$${BASENAME}_organized.html"; \
	else \
		OUTPUT_FILE="$(OUTPUT_DIR)/$(OUTPUT)"; \
	fi; \
	EXTRA_CSV=""; \
	if [ -n "$(SHIP)" ]; then EXTRA_CSV=$$(eval echo "$(SHIP)"); fi; \
	$(PYTHON) mtg_packing_slip_organizer.py "$$EXPANDED_CSV" $${EXTRA_CSV:+"$$EXTRA_CSV"} -o "$$OUTPUT_FILE"

serve: ## Run the web app locally
	$(VENV)/bin/uvicorn app:app --reload --port $(PORT)

test: ## Run the tests
	$(PYTHON) -m pytest -q tests

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
