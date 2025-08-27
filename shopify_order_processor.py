# -*- coding: utf-8 -*-
"""
Shopify Order Processor

This script processes Shopify order exports (CSV) to filter, aggregate,
and format them into a clean Excel report.
"""
import sys
import os
from datetime import datetime

import pandas as pd
# openpyxl is used by pandas for Excel writing, so we import it to ensure it's available.
import openpyxl

def validate_date_format(date_string):
    """Validates that the date string is in DD.MM.YYYY format and returns a datetime object or None."""
    try:
        return datetime.strptime(date_string, "%d.%m.%Y")
    except ValueError:
        return None

def get_tariff_from_user(prompt_message):
    """Prompts the user for a tariff value and validates it in a loop."""
    while True:
        cost_str = input(prompt_message)
        try:
            cost = float(cost_str)
            if cost >= 0:
                return cost
            else:
                print("Error: Please enter a non-negative number.", file=sys.stderr)
        except ValueError:
            print("Error: Invalid input. Please enter a valid number (e.g., 10.50).", file=sys.stderr)

def get_date_from_user(prompt_message):
    """Prompts the user for a date, validates it, and loops until a valid date is entered."""
    while True:
        date_str = input(prompt_message)
        date_obj = validate_date_format(date_str)
        if date_obj:
            return date_obj
        else:
            print(f"Error: Invalid date format. Please use DD.MM.YYYY.", file=sys.stderr)

def load_and_validate_csv(file_path):
    """Loads the CSV, checks for file existence, and validates required columns."""
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' was not found.", file=sys.stderr)
        sys.exit(1)

    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error: Failed to read the CSV file. Reason: {e}", file=sys.stderr)
        sys.exit(1)

    required_columns = [
        'Name', 'Fulfilled at', 'Lineitem quantity',
        'Lineitem name', 'Lineitem sku', 'Total'
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f"Error: The input CSV is missing the following required columns: {', '.join(missing_columns)}", file=sys.stderr)
        sys.exit(1)

    print("Successfully loaded and validated the input file.")
    return df

def filter_by_date_range(df, start_date, end_date):
    """Filters the DataFrame based on the 'Fulfilled at' date column."""
    print(f"Initial record count: {len(df)}")

    # Forward-fill key columns to propagate their values to all line items of an order.
    ffill_cols = ['Fulfilled at', 'Fulfillment Status', 'Financial Status']
    for col in ffill_cols:
        if col in df.columns:
            df[col] = df.groupby('Name')[col].ffill()

    # Drop rows with no fulfillment date after f-fill (unfulfilled orders)
    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        print(f"Dropped {original_count - len(df)} unfulfilled orders.")

    # Convert 'Fulfilled at' to datetime objects.
    df['Fulfilled at'] = pd.to_datetime(df['Fulfilled at'], errors='coerce')

    # Drop rows that could not be parsed into a valid date.
    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        print(f"Dropped {original_count - len(df)} rows with invalid date format in 'Fulfilled at'.")

    # After parsing, make the column timezone-naive to allow comparison.
    if pd.api.types.is_datetime64_any_dtype(df['Fulfilled at']) and df['Fulfilled at'].dt.tz is not None:
        df['Fulfilled at'] = df['Fulfilled at'].dt.tz_localize(None)

    # Filter the DataFrame to include only orders within the specified date range.
    mask = (df['Fulfilled at'].dt.normalize() >= start_date) & (df['Fulfilled at'].dt.normalize() <= end_date)
    filtered_df = df.loc[mask].copy()

    print(f"Record count after filtering by date: {len(filtered_df)}")

    if len(filtered_df) == 0:
        print("Warning: No orders found within the specified date range.")

    return filtered_df

def calculate_costs(df, cost_first_sku, cost_next_sku, cost_per_piece):
    """Calculates the processing cost for each order based on tariffs."""
    if df.empty:
        return df

    # Exclude "Package protection" items for billing calculation
    billable_df = df[~df['Lineitem name'].str.contains("Package protection", na=False)].copy()

    if billable_df.empty:
        # If no billable items, return original df with zero-cost columns
        df['Billable Unique SKUs'] = 0
        df['Billable Total Quantity'] = 0
        df['SKU Cost'] = 0
        df['Quantity Cost'] = 0
        df['Total Order Cost'] = 0
        return df

    # Group by order to get SKU counts and quantities from billable items
    order_summary = billable_df.groupby('Name').agg(
        Billable_Unique_SKUs=('Lineitem sku', 'nunique'),
        Billable_Total_Quantity=('Lineitem quantity', 'sum')
    ).reset_index()

    # Calculate costs for each order
    def calculate_sku_cost(sku_count):
        if sku_count == 0:
            return 0
        return cost_first_sku + (max(0, sku_count - 1) * cost_next_sku)

    order_summary['SKU Cost'] = order_summary['Billable_Unique_SKUs'].apply(calculate_sku_cost)
    order_summary['Quantity Cost'] = order_summary['Billable_Total_Quantity'] * cost_per_piece
    order_summary['Total Order Cost'] = order_summary['SKU Cost'] + order_summary['Quantity Cost']

    # Merge the calculated costs back into the original DataFrame
    df_with_costs = pd.merge(df, order_summary, on='Name', how='left')

    # Fill NaN for cost columns with 0 (for orders that had no billable items)
    cost_cols = ['Billable_Unique_SKUs', 'Billable_Total_Quantity', 'SKU Cost', 'Quantity Cost', 'Total Order Cost']
    for col in cost_cols:
        if col in df_with_costs.columns:
            df_with_costs[col].fillna(0, inplace=True)

    # Ensure integer columns are of integer type
    int_cols = ['Billable_Unique_SKUs', 'Billable_Total_Quantity']
    for col in int_cols:
        if col in df_with_costs.columns:
            df_with_costs[col] = df_with_costs[col].astype(int)

    return df_with_costs

def create_invoice_summary(df_with_costs, cost_first_sku, cost_next_sku, cost_per_piece):
    """Creates a DataFrame with a summary of all costs for the invoice."""
    if df_with_costs.empty:
        return pd.DataFrame()

    # Get one row per order to sum up order-level costs without duplication
    order_costs = df_with_costs.drop_duplicates(subset=['Name'])

    total_orders = order_costs['Name'].nunique()
    if total_orders == 0:
        return pd.DataFrame()

    total_sku_cost = order_costs['SKU Cost'].sum()
    total_quantity_cost = order_costs['Quantity Cost'].sum()
    grand_total_cost = order_costs['Total Order Cost'].sum()

    # Calculate total counts of first SKUs and next SKUs
    total_billable_skus = order_costs['Billable_Unique_SKUs'].sum()
    total_first_skus = total_orders if total_billable_skus > 0 else 0
    total_next_skus = total_billable_skus - total_first_skus

    total_pieces = order_costs['Billable_Total_Quantity'].sum()

    # Create the summary DataFrame
    summary_data = {
        'Description': [
            'Total Orders Processed',
            'First SKU Tariff',
            'Subsequent SKU Tariff',
            'Per-Piece Tariff',
            '---',
            'Total SKU Cost',
            'Total Piece Cost',
            '---',
            'GRAND TOTAL'
        ],
        'Rate': [
            '',
            f'{cost_first_sku:.2f}',
            f'{cost_next_sku:.2f}',
            f'{cost_per_piece:.2f}',
            '',
            '',
            '',
            '',
            ''
        ],
        'Count': [
            total_orders,
            total_first_skus,
            total_next_skus,
            total_pieces,
            '',
            '',
            '',
            '',
            ''
        ],
        'Total Amount': [
            '',
            total_first_skus * cost_first_sku,
            total_next_skus * cost_next_sku,
            total_pieces * cost_per_piece,
            '',
            f'{total_sku_cost:.2f}',
            f'{total_quantity_cost:.2f}',
            '',
            f'{grand_total_cost:.2f}'
        ]
    }

    summary_df = pd.DataFrame(summary_data)
    return summary_df

from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_excel_report(sheets_data, output_filename):
    """Formats and saves one or more DataFrames to a multi-sheet Excel file with custom borders."""
    if not sheets_data:
        print("No data to save, skipping Excel report generation.")
        return

    # Define border styles
    thin_side = Side(border_style="thin", color="000000")
    thick_side = Side(border_style="thick", color="000000")

    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    thick_top_border = Border(top=thick_side)
    thick_bottom_border = Border(bottom=thick_side)
    left_thick_border = Border(left=thick_side)
    right_thick_border = Border(right=thick_side)

    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            for sheet_name, df in sheets_data.items():
                if df.empty:
                    continue
                df.to_excel(writer, index=False, sheet_name=sheet_name)

            for sheet_name, worksheet in writer.sheets.items():
                df = sheets_data[sheet_name]

                # --- Basic Formatting (apply to all sheets) ---
                header_font = Font(bold=True)
                for cell in worksheet["1:1"]:
                    cell.font = header_font
                    cell.border = thin_border # Add border to header

                worksheet.freeze_panes = 'A2'
                if df.shape[0] > 0: # Only add filter if there is data
                    worksheet.auto_filter.ref = worksheet.dimensions

                # --- Column Width Formatting (apply to all sheets) ---
                for i, col in enumerate(df.columns, 1):
                    column_letter = get_column_letter(i)
                    if col == 'Fulfilled at':
                        worksheet.column_dimensions[column_letter].width = 20
                        for cell in worksheet[column_letter][1:]:
                            if cell.value: cell.number_format = 'DD.MM.YYYY HH:MM'
                    else:
                        if not df[col].empty: max_length = max(df[col].astype(str).map(len).max(), len(col))
                        else: max_length = len(col)
                        worksheet.column_dimensions[column_letter].width = max_length + 2

                # --- Advanced Border Formatting (conditional) ---
                if sheet_name in ['All Orders', 'Without Package Protection', 'Cost Calculation']:
                    # Group by order number to identify row groups
                    order_groups = df.groupby((df['Name'] != df['Name'].shift()).cumsum())

                    for _, group in order_groups:
                        min_row = group.index.min() + 2  # +2 for header and 0-indexing
                        max_row = group.index.max() + 2

                        for row_idx in range(min_row, max_row + 1):
                            for col_idx in range(1, len(df.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)

                                # Determine the border sides for the current cell
                                top = thick_side if row_idx == min_row else thin_side
                                bottom = thick_side if row_idx == max_row else thin_side
                                left = thick_side if col_idx == 1 else thin_side
                                right = thick_side if col_idx == len(df.columns) else thin_side

                                # For rows in the middle of a multi-row group, the top border should be thin.
                                if row_idx > min_row:
                                    top = thin_side

                                cell.border = Border(left=left, right=right, top=top, bottom=bottom)

                elif sheet_name == 'Final Invoice':
                    # Apply a simple border to the entire table
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, max_col=worksheet.max_column):
                        for cell in row:
                            cell.border = thin_border

        print(f"Successfully created Excel report: {output_filename}")

    except Exception as e:
        print(f"Error: Failed to create Excel report. Reason: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """Main function to run the script."""
    print("Starting Shopify Order Processor...")

    # Get start and end dates from user
    start_date = get_date_from_user("Enter the start date (DD.MM.YYYY): ")
    end_date = get_date_from_user("Enter the end date (DD.MM.YYYY): ")

    # Get tariffs from user
    print("\nPlease enter the cost tariffs:")
    cost_first_sku = get_tariff_from_user("Enter the cost for the first SKU: ")
    cost_next_sku = get_tariff_from_user("Enter the cost for each subsequent SKU: ")
    cost_per_piece = get_tariff_from_user("Enter the cost per piece: ")

    input_filename = "orders_export.csv"
    print(f"\nInput file: {input_filename}")
    print(f"Processing orders from {start_date.strftime('%d.%m.%Y')} to {end_date.strftime('%d.%m.%Y')}")

    # 1. Load and validate the CSV file
    source_df = load_and_validate_csv(input_filename)

    # 2. Filter orders by date range
    filtered_df = filter_by_date_range(source_df, start_date, end_date)

    # 3. Select and format final columns for the report
    final_columns = [
        'Name', 'Fulfilled at', 'Fulfillment Status', 'Financial Status',
        'Created at', 'Lineitem quantity', 'Lineitem name', 'Lineitem sku',
        'Lineitem fulfillment status'
    ]
    # Select only the columns that actually exist in the dataframe
    existing_columns = [col for col in final_columns if col in filtered_df.columns]
    report_df = filtered_df[existing_columns]

    # 4. Create DataFrames for Excel sheets
    if not report_df.empty:
        # Create a second DataFrame excluding "Package protection"
        df_no_protection = report_df[~report_df['Lineitem name'].str.contains("Package protection", na=False)].copy()

        # Create the detailed cost calculation DataFrame (Sheet 3)
        df_with_costs = calculate_costs(report_df, cost_first_sku, cost_next_sku, cost_per_piece)

        # Create the final invoice summary DataFrame (Sheet 4)
        df_invoice = create_invoice_summary(df_with_costs, cost_first_sku, cost_next_sku, cost_per_piece)

        sheets_data = {
            'All Orders': report_df,
            'Without Package Protection': df_no_protection,
            'Cost Calculation': df_with_costs,
            'Final Invoice': df_invoice
        }

        # Prompt user for output filename
        prompt_message = "Enter the desired name for the output Excel file (e.g., report.xlsx).\nPress Enter to use a default name: "
        output_filename_from_user = input(prompt_message)

        if output_filename_from_user:
            output_filename = output_filename_from_user
            # Ensure the filename has .xlsx extension
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
        else:
            # Generate a dynamic filename if not provided
            current_date = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"processed_orders_{current_date}.xlsx"
            print(f"No filename provided. Using default: {output_filename}")

        create_excel_report(sheets_data, output_filename)
    else:
        print("Script finished. No orders to process into an Excel file.")

if __name__ == "__main__":
    main()
