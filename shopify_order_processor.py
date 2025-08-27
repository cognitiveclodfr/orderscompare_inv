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
from tqdm import tqdm

def validate_date_format(date_string):
    """Validates that the date string is in DD.MM.YYYY format and returns a datetime object or None."""
    try:
        return datetime.strptime(date_string, "%d.%m.%Y")
    except ValueError:
        return None

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

    # Forward-fill 'Fulfilled at' to propagate the date to all line items of an order.
    df['Fulfilled at'] = df.groupby('Name')['Fulfilled at'].ffill()

    # Drop rows with no fulfillment date, as they are not relevant for the report.
    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        print(f"Dropped {original_count - len(df)} rows with empty 'Fulfilled at' date.")

    # Convert 'Fulfilled at' to datetime objects. Errors will be coerced to NaT (Not a Time).
    # Based on user feedback, we do not perform timezone conversion.
    df['Fulfilled at'] = pd.to_datetime(df['Fulfilled at'], errors='coerce')

    # After parsing, make the column timezone-naive to allow comparison.
    if pd.api.types.is_datetime64_any_dtype(df['Fulfilled at']) and df['Fulfilled at'].dt.tz is not None:
        df['Fulfilled at'] = df['Fulfilled at'].dt.tz_localize(None)

    # Drop rows that could not be parsed into a valid date.
    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        print(f"Dropped {original_count - len(df)} rows with invalid date format in 'Fulfilled at'.")

    # Filter the DataFrame to include only orders within the specified date range.
    # We normalize the date part to ensure the comparison is inclusive of the whole day.
    mask = (df['Fulfilled at'].dt.normalize() >= start_date) & (df['Fulfilled at'].dt.normalize() <= end_date)
    filtered_df = df.loc[mask]

    print(f"Record count after filtering by date: {len(filtered_df)}")

    if len(filtered_df) == 0:
        print("Warning: No orders found within the specified date range.")

    return filtered_df

def aggregate_orders(df):
    """Groups order line items into single orders and aggregates the data."""
    if df.empty:
        print("No data to aggregate.")
        return pd.DataFrame()

    # Configure tqdm for pandas
    tqdm.pandas(desc="Aggregating Orders")

    # Define aggregation functions
    # Using a lambda with a check for all-NaN to avoid warnings
    def join_unique(x):
        return '\n'.join(x.dropna().astype(str).unique())

    aggregations = {
        'Fulfilled at': lambda x: x.iloc[0],
        'Lineitem name': lambda x: join_unique(x),
        'Lineitem quantity': 'sum',
        'Lineitem sku': lambda x: join_unique(x),
        'Total': lambda x: x.iloc[0]
    }

    # Group by order name and apply aggregations with a progress bar
    aggregated_df = df.groupby('Name').progress_apply(lambda x: x.agg(aggregations)).reset_index()

    # Calculate the number of unique items (positions)
    # This needs to be done separately as it requires a nunique on the original group
    unique_items_count = df.groupby('Name')['Lineitem name'].nunique().reset_index(name='Unique Items')

    # Merge the unique items count back into the aggregated dataframe
    aggregated_df = pd.merge(aggregated_df, unique_items_count, on='Name')

    # Rename columns to the final English names
    aggregated_df.rename(columns={
        'Name': 'Order Number',
        'Fulfilled at': 'Fulfillment Date',
        'Unique Items': 'Unique Items',
        'Lineitem quantity': 'Total Quantity',
        'Lineitem name': 'Article List',
        'Lineitem sku': 'Article SKUs',
        'Total': 'Grand Total'
    }, inplace=True)

    # Reorder columns to the desired output format
    final_columns = [
        'Order Number', 'Fulfillment Date', 'Unique Items', 'Total Quantity',
        'Article List', 'Article SKUs', 'Grand Total'
    ]
    aggregated_df = aggregated_df[final_columns]

    print(f"Successfully aggregated {len(aggregated_df)} orders.")

    return aggregated_df

from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

def create_excel_report(df, output_filename):
    """Formats and saves the final DataFrame to an Excel file."""
    if df.empty:
        print("No data to save, skipping Excel report generation.")
        return

    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Processed Orders')

            # Access the workbook and worksheet objects for formatting
            workbook = writer.book
            worksheet = writer.sheets['Processed Orders']

            # 1. Apply formatting: Bold headers, freeze top row, add filters
            header_font = Font(bold=True)
            for cell in worksheet["1:1"]:
                cell.font = header_font

            worksheet.freeze_panes = 'A2'
            worksheet.auto_filter.ref = worksheet.dimensions

            # 2. Auto-fit column widths and apply specific formats
            for i, col in enumerate(df.columns, 1):
                column_letter = get_column_letter(i)

                if col == 'Fulfilled at':
                    # Apply date format and set width
                    worksheet.column_dimensions[column_letter].width = 20
                    for cell in worksheet[column_letter][1:]:
                        if cell.value:
                            cell.number_format = 'DD.MM.YYYY HH:MM'
                else:
                    # Auto-fit other columns
                    if not df[col].empty:
                        max_length = max(df[col].astype(str).map(len).max(), len(col))
                    else:
                        max_length = len(col)
                    worksheet.column_dimensions[column_letter].width = max_length + 2

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

    input_filename = "orders_export.csv"
    print(f"\nInput file: {input_filename}")
    print(f"Processing orders from {start_date.strftime('%d.%m.%Y')} to {end_date.strftime('%d.%m.%Y')}")

    # 1. Load and validate the CSV file
    source_df = load_and_validate_csv(input_filename)

    # 2. Filter orders by date range
    filtered_df = filter_by_date_range(source_df, start_date, end_date)

    # 3. Aggregate order data
    aggregated_df = aggregate_orders(filtered_df)

    # 4. Create the Excel report
    if not aggregated_df.empty:
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

        create_excel_report(aggregated_df, output_filename)
    else:
        print("Script finished. No orders to process into an Excel file.")

if __name__ == "__main__":
    main()
