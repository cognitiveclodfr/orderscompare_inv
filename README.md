# Shopify Order Processor

This script is a command-line tool for processing Shopify order export CSV files. It filters orders by a specified date range, aggregates the data to provide one summary row per order, and saves the result in a well-formatted Excel file.

## Features

- **Filter by Date:** Process only the orders fulfilled within a specific date range.
- **Data Aggregation:** Converts raw, line-item-based exports into a clean, order-level summary.
- **Calculated Fields:** Automatically calculates total items, unique items, and creates newline-separated lists of products and SKUs for each order.
- **Formatted Excel Output:** Generates a professional Excel report with bold headers, auto-fitted columns, filters, a frozen header row, and specific formatting for dates and lists.
- **Progress Bar:** Displays a progress bar using `tqdm` when processing large files.
- **Error Handling:** Provides clear error messages for common issues like missing files or incorrect data formats.

## Prerequisites

- Python 3.6 or higher

## Installation

1.  **Clone the repository or download the source code.**

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Run the script from your terminal using the following command structure. You must provide the path to the input CSV file, a start date, and an end date.

### Syntax
```bash
python shopify_order_processor.py <path_to_csv> --start-date <DD.MM.YYYY> --end-date <DD.MM.YYYY> [--output <output_filename.xlsx>]
```

### Arguments
- `filepath`: (Required) The full path to the Shopify orders CSV file.
- `--start-date`: (Required) The start of the reporting period in `DD.MM.YYYY` format.
- `--end-date`: (Required) The end of the reporting period in `DD.MM.YYYY` format.
- `--output`: (Optional) The desired name for the output Excel file. If not provided, a default name like `processed_orders_YYYY-MM-DD.xlsx` will be used.

### Example
```bash
python shopify_order_processor.py "downloads/orders_export_2025.csv" --start-date "01.07.2025" --end-date "31.07.2025" --output "july_2025_report.xlsx"
```
