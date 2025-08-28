# Shopify Order Processor

This script is an interactive tool for processing Shopify order export CSV files. It filters orders by a specified date range, calculates processing costs based on user-defined tariffs, and saves the result in a detailed, multi-sheet Excel file.

## Features

- **Interactive Prompts**: Guides the user to enter the required date range and cost tariffs directly in the terminal.
- **Robust Data Handling**: Automatically cleans order numbers (e.g., removing extra whitespace) to prevent grouping errors with inconsistent data.
- **Date Filtering**: Processes only the orders fulfilled within the specified start and end dates.
- **Line-Item Cost Calculation**:
    - Calculates processing cost for **each line item** based on three user-defined tariffs:
        1.  A flat cost for the **first billable SKU** in an order.
        2.  A cost for each **subsequent unique SKU**.
        3.  A cost for each **individual item (piece)**.
    - Excludes "Package protection" and "Shipping Protection" line items from all cost calculations.
- **Multi-Sheet Excel Reports**: Generates a professionally formatted Excel report with four distinct sheets:
    1.  **All Orders**: A detailed view of all filtered order line items.
    2.  **Without Package Protection**: The same data as the first sheet but excludes protection line items.
    3.  **Cost Calculation**: A transparent breakdown of the calculated costs for each order, with costs detailed per line item and a `TOTAL` row summarizing each order.
    4.  **Final Invoice**: A high-level summary sheet that aggregates all costs into a final invoice format, perfect for billing.
- **Advanced Excel Formatting**: The output file includes:
    - Bold headers and a frozen header row.
    - Auto-fitted column widths.
    - Table filters enabled on all sheets.
    - Custom borders to visually group line items, with a thick border under each order's `TOTAL` row for clarity.
- **Error Handling**: Provides clear error messages for common issues like a missing input file or incorrect data formats.

## Prerequisites

- Python 3.6 or higher
- An `orders_export.csv` file from Shopify in the same directory as the script.

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

This script is run interactively. No command-line arguments are needed.

1.  **Place your data file** in the same directory as the script and ensure it is named `orders_export.csv`.

2.  **Run the script** from your terminal:
    ```bash
    python shopify_order_processor.py
    ```

3.  **Follow the prompts** to enter the start date, end date, and the three cost tariffs.

    ```
    Starting Shopify Order Processor...
    Enter the start date (DD.MM.YYYY): 01.07.2025
    Enter the end date (DD.MM.YYYY): 31.07.2025

    Please enter the cost tariffs:
    Enter the cost for the first SKU: 1.50
    Enter the cost for each subsequent SKU: 0.75
    Enter the cost per piece: 0.25
    ```

4.  **Provide an output filename** when prompted, or press Enter to use a default name (e.g., `processed_orders_YYYY-MM-DD.xlsx`).

5.  Once the script finishes, the Excel report will be saved in the same directory.

## Output File Structure

The generated Excel file contains the following sheets:

- **All Orders**: Shows every line item for all orders that were fulfilled within the specified date range.
- **Without Package Protection**: A filtered version of the "All Orders" sheet, which is useful if you want to see order details without the protection items.
- **Cost Calculation**: This sheet provides a transparent, line-by-line breakdown of costs for each order.
    - It adds new columns: `Piece Cost`, `SKU Cost`, and `Line Total Cost`.
    - Each billable line item has its costs calculated and displayed in its own row.
    - At the end of each order's items, a **`TOTAL`** row is added, summing up the costs for that specific order. This row is clearly marked with a thick border underneath.
- **Final Invoice**: A clean, high-level summary of the total costs across all processed orders. It breaks down the total charges by tariff, providing a clear and simple invoice for billing purposes. The totals are derived from the new line-item calculations.
