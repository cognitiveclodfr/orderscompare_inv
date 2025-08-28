# -*- coding: utf-8 -*-
"""
Shopify Order Processor

This script processes Shopify order exports (CSV) to filter, aggregate,
and format them into a clean Excel report.
"""
import sys
import os
import logging
from datetime import datetime

import pandas as pd
# openpyxl is used by pandas for Excel writing, so we import it to ensure it's available.
import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter

def validate_date_format(date_string):
    """
    Validates that a date string is in the DD.MM.YYYY format.

    Args:
        date_string (str): The string to validate.

    Returns:
        datetime.datetime | None: A datetime object if the format is correct,
                                  otherwise None.
    """
    try:
        return datetime.strptime(date_string, "%d.%m.%Y")
    except ValueError:
        return None

def get_tariff_from_user(prompt_message):
    """
    Prompts the user for a numeric tariff value and validates it.

    The function will loop indefinitely until the user enters a non-negative
    number.

    Args:
        prompt_message (str): The message to display to the user.

    Returns:
        float: The validated, non-negative tariff value.
    """
    while True:
        cost_str = input(prompt_message)
        try:
            cost = float(cost_str)
            if cost >= 0:
                return cost
            else:
                logging.error("Please enter a non-negative number.")
        except ValueError:
            logging.error("Invalid input. Please enter a valid number (e.g., 10.50).")

def get_date_from_user(prompt_message):
    """
    Prompts the user for a date and validates it is in DD.MM.YYYY format.

    The function will loop indefinitely until the user enters a valid date
    string.

    Args:
        prompt_message (str): The message to display to the user.

    Returns:
        datetime.datetime: The validated date as a datetime object.
    """
    while True:
        date_str = input(prompt_message)
        date_obj = validate_date_format(date_str)
        if date_obj:
            return date_obj
        else:
            logging.error("Invalid date format. Please use DD.MM.YYYY.")

def load_and_validate_csv(file_path):
    """
    Loads a CSV file, validates its existence and required columns.

    It specifies data types for certain columns to ensure efficient loading
    and prevent type errors. If the file is not found or is missing
    essential columns, the script will exit.

    Args:
        file_path (str): The path to the input CSV file.

    Returns:
        pd.DataFrame: The loaded and validated DataFrame.

    Exits:
        The script will exit with status 1 if the file is not found, cannot be
        read, or is missing required columns.
    """
    if not os.path.exists(file_path):
        logging.error(f"The file '{file_path}' was not found.")
        sys.exit(1)

    dtype_map = {
        'Subtotal': float, 'Shipping': float, 'Taxes': float, 'Total': float,
        'Discount Amount': float, 'Lineitem quantity': int, 'Lineitem price': float,
        'Lineitem compare at price': float, 'Lineitem requires shipping': bool,
        'Lineitem taxable': bool, 'Refunded Amount': float, 'Outstanding Balance': float,
        'Id': float, 'Lineitem discount': float, 'Tax 1 Value': float, 'Tax 2 Value': float,
        'Tax 3 Value': float, 'Tax 4 Value': float, 'Tax 5 Value': float, 'Duties': float,
    }

    try:
        df = pd.read_csv(file_path, dtype=dtype_map, low_memory=False)
    except Exception as e:
        logging.error(f"Failed to read the CSV file. Reason: {e}")
        sys.exit(1)

    required_columns = [
        'Name', 'Fulfilled at', 'Lineitem quantity',
        'Lineitem name', 'Lineitem sku', 'Total'
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logging.error(f"The input CSV is missing the following required columns: {', '.join(missing_columns)}")
        sys.exit(1)

    logging.info("Successfully loaded and validated the input file.")
    return df

def filter_by_date_range(df, start_date, end_date):
    """
    Filters a DataFrame of orders by a given date range.

    The filtering is based on the 'Fulfilled at' column. The function handles
    missing fulfillment dates, date parsing errors, and timezone conversion.

    Args:
        df (pd.DataFrame): The input DataFrame of orders.
        start_date (datetime.datetime): The start of the filtering period.
        end_date (datetime.datetime): The end of the filtering period.

    Returns:
        pd.DataFrame: A new DataFrame containing only the orders within the
                      specified date range.
    """
    logging.info(f"Initial record count: {len(df)}")

    ffill_cols = ['Fulfilled at', 'Fulfillment Status', 'Financial Status']
    for col in ffill_cols:
        if col in df.columns:
            df[col] = df.groupby('Name')[col].ffill()

    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        logging.info(f"Dropped {original_count - len(df)} unfulfilled orders.")

    df['Fulfilled at'] = pd.to_datetime(df['Fulfilled at'], errors='coerce')

    original_count = len(df)
    df.dropna(subset=['Fulfilled at'], inplace=True)
    if len(df) < original_count:
        logging.warning(f"Dropped {original_count - len(df)} rows with invalid date format in 'Fulfilled at'.")

    if pd.api.types.is_datetime64_any_dtype(df['Fulfilled at']) and df['Fulfilled at'].dt.tz is not None:
        df['Fulfilled at'] = df['Fulfilled at'].dt.tz_localize(None)

    mask = (df['Fulfilled at'].dt.normalize() >= start_date) & (df['Fulfilled at'].dt.normalize() <= end_date)
    filtered_df = df.loc[mask].copy()

    logging.info(f"Record count after filtering by date: {len(filtered_df)}")
    if len(filtered_df) == 0:
        logging.warning("No orders found within the specified date range.")
    return filtered_df

def calculate_costs(df, cost_first_sku, cost_next_sku, cost_per_piece):
    """
    Calculates processing costs for each order and adds them to the DataFrame.
    """
    if df.empty:
        return df

    billable_df = df[~df['Lineitem name'].str.contains("Package protection", na=False)].copy()
    if billable_df.empty:
        df['Billable_Unique_SKUs'] = 0
        df['Billable_Total_Quantity'] = 0
        df['SKU Cost'] = 0.0
        df['Quantity Cost'] = 0.0
        df['Total Order Cost'] = 0.0
        return df

    order_summary = billable_df.groupby('Name').agg(
        Billable_Unique_SKUs=('Lineitem sku', 'nunique'),
        Billable_Total_Quantity=('Lineitem quantity', 'sum')
    ).reset_index()

    def calculate_sku_cost(sku_count):
        if sku_count == 0:
            return 0
        return cost_first_sku + (max(0, sku_count - 1) * cost_next_sku)

    order_summary['SKU Cost'] = order_summary['Billable_Unique_SKUs'].apply(calculate_sku_cost)
    order_summary['Quantity Cost'] = order_summary['Billable_Total_Quantity'] * cost_per_piece
    order_summary['Total Order Cost'] = order_summary['SKU Cost'] + order_summary['Quantity Cost']

    df_with_costs = pd.merge(df, order_summary, on='Name', how='left')
    cost_cols = ['Billable_Unique_SKUs', 'Billable_Total_Quantity', 'SKU Cost', 'Quantity Cost', 'Total Order Cost']
    df_with_costs[cost_cols] = df_with_costs[cost_cols].fillna(0)
    int_cols = ['Billable_Unique_SKUs', 'Billable_Total_Quantity']
    df_with_costs[int_cols] = df_with_costs[int_cols].astype(int)

    return df_with_costs

def create_invoice_summary(df_with_costs, cost_first_sku, cost_next_sku, cost_per_piece):
    """
    Creates a summary DataFrame formatted as a final invoice.
    """
    if df_with_costs.empty:
        return pd.DataFrame()

    order_costs = df_with_costs.drop_duplicates(subset=['Name'])
    total_orders = order_costs['Name'].nunique()
    if total_orders == 0:
        return pd.DataFrame()

    total_sku_cost = order_costs['SKU Cost'].sum()
    total_quantity_cost = order_costs['Quantity Cost'].sum()
    grand_total_cost = order_costs['Total Order Cost'].sum()
    total_billable_skus = order_costs['Billable_Unique_SKUs'].sum()
    total_first_skus = total_orders if total_billable_skus > 0 else 0
    total_next_skus = total_billable_skus - total_first_skus
    total_pieces = order_costs['Billable_Total_Quantity'].sum()

    summary_data = {
        'Description': [
            'Total Orders Processed', 'First SKU Tariff', 'Subsequent SKU Tariff', 'Per-Piece Tariff',
            '---', 'Total SKU Cost', 'Total Piece Cost', '---', 'GRAND TOTAL'
        ],
        'Rate': [
            '', f'{cost_first_sku:.2f}', f'{cost_next_sku:.2f}', f'{cost_per_piece:.2f}',
            '', '', '', '', ''
        ],
        'Count': [
            total_orders, total_first_skus, total_next_skus, total_pieces,
            '', '', '', '', ''
        ],
        'Total Amount': [
            '', total_first_skus * cost_first_sku, total_next_skus * cost_next_sku, total_pieces * cost_per_piece,
            '', f'{total_sku_cost:.2f}', f'{total_quantity_cost:.2f}', '', f'{grand_total_cost:.2f}'
        ]
    }
    return pd.DataFrame(summary_data)

def transform_cost_df_for_reporting(df_costs):
    """
    Transforms the cost calculation DataFrame to include a summary 'TOTAL' row for each order.

    Args:
        df_costs (pd.DataFrame): The DataFrame from the `calculate_costs` function.

    Returns:
        pd.DataFrame: A new DataFrame with a 'TOTAL' row for each order, ready for reporting.
    """
    if df_costs.empty:
        return df_costs

    processed_orders = []
    for name, group in df_costs.groupby('Name'):
        # Calculate totals for the group
        total_quantity = group['Lineitem quantity'].sum()
        total_sku_cost = group['SKU Cost'].sum()
        total_quantity_cost = group['Quantity Cost'].sum()
        # Total order cost is the same for all rows in the group
        total_order_cost = group['Total Order Cost'].iloc[0]

        # Create the 'TOTAL' row
        total_row = pd.DataFrame([{
            'Name': name,
            'Lineitem name': 'TOTAL',
            'Lineitem quantity': total_quantity,
            'SKU Cost': total_sku_cost,
            'Quantity Cost': total_quantity_cost,
            'Total Order Cost': total_order_cost
        }])

        # Clear the 'Total Order Cost' from the individual item rows
        group_copy = group.copy()
        group_copy['Total Order Cost'] = pd.NA

        # Combine the original group with the new total row
        processed_group = pd.concat([group_copy, total_row], ignore_index=True)
        processed_orders.append(processed_group)

    # Combine all processed groups back into a single DataFrame
    final_df = pd.concat(processed_orders, ignore_index=True)
    return final_df

def create_excel_report(sheets_data, output_filename):
    """
    Creates a multi-sheet Excel report from a dictionary of DataFrames.
    """
    if not sheets_data:
        logging.info("No data to save, skipping Excel report generation.")
        return

    thin_side = Side(border_style="thin", color="000000")
    thick_side = Side(border_style="thick", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            for sheet_name, df in sheets_data.items():
                if df.empty: continue
                if sheet_name in ['All Orders', 'Without Package Protection', 'Cost Calculation']:
                    df = df.sort_values(by='Name').reset_index(drop=True)
                df.to_excel(writer, index=False, sheet_name=sheet_name)

            for sheet_name, worksheet in writer.sheets.items():
                if worksheet.max_row <= 1: continue

                header_font = Font(bold=True)
                for cell in worksheet["1:1"]:
                    cell.font = header_font
                    cell.border = thin_border
                worksheet.freeze_panes = 'A2'
                worksheet.auto_filter.ref = worksheet.dimensions

                df_to_format = sheets_data[sheet_name]
                if sheet_name in ['All Orders', 'Without Package Protection', 'Cost Calculation']:
                     df_to_format = df_to_format.sort_values(by='Name').reset_index(drop=True)

                for i, col_name in enumerate(df_to_format.columns, 1):
                    column_letter = get_column_letter(i)
                    if col_name == 'Fulfilled at':
                        worksheet.column_dimensions[column_letter].width = 20
                        for cell in worksheet[column_letter][1:]:
                            if cell.value: cell.number_format = 'DD.MM.YYYY HH:MM'
                    else:
                        max_len = max((df_to_format[col_name].astype(str).map(len).max(), len(col_name)))
                        worksheet.column_dimensions[column_letter].width = max_len + 2

                if sheet_name in ['All Orders', 'Without Package Protection', 'Cost Calculation']:
                    thin_bottom_border = Border(bottom=thin_side)
                    thick_bottom_border = Border(bottom=thick_side)
                    for row_idx in range(2, worksheet.max_row + 1):
                        is_last = (row_idx == worksheet.max_row or
                                   worksheet.cell(row=row_idx, column=1).value != worksheet.cell(row=row_idx + 1, column=1).value)
                        border_to_apply = thick_bottom_border if is_last else thin_bottom_border
                        for col_idx in range(1, worksheet.max_column + 1):
                            worksheet.cell(row=row_idx, column=col_idx).border = border_to_apply
                elif sheet_name == 'Final Invoice':
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                        for cell in row:
                            cell.border = thin_border

        logging.info(f"Successfully created Excel report: {output_filename}")
    except Exception as e:
        logging.error(f"Failed to create Excel report. Reason: {e}")
        sys.exit(1)

def prepare_report_sheets(report_df, cost_first_sku, cost_next_sku, cost_per_piece):
    """
    Prepares a dictionary of DataFrames for the multi-sheet Excel report.
    """
    df_no_protection = report_df[~report_df['Lineitem name'].str.contains("Package protection", na=False)].copy()
    df_with_costs = calculate_costs(report_df, cost_first_sku, cost_next_sku, cost_per_piece)

    # Transform the cost calculation sheet to have the new summary format
    df_costs_transformed = transform_cost_df_for_reporting(df_with_costs)

    df_invoice = create_invoice_summary(df_with_costs, cost_first_sku, cost_next_sku, cost_per_piece)
    return {
        'All Orders': report_df,
        'Without Package Protection': df_no_protection,
        'Cost Calculation': df_costs_transformed,
        'Final Invoice': df_invoice
    }

def main():
    """
    Main function to orchestrate the order processing workflow.
    """
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info("Starting Shopify Order Processor...")

    # --- User Input ---
    start_date = get_date_from_user("Enter the start date (DD.MM.YYYY): ")
    end_date = get_date_from_user("Enter the end date (DD.MM.YYYY): ")
    print("\nPlease enter the cost tariffs:")
    cost_first_sku = get_tariff_from_user("Enter the cost for the first SKU: ")
    cost_next_sku = get_tariff_from_user("Enter the cost for each subsequent SKU: ")
    cost_per_piece = get_tariff_from_user("Enter the cost per piece: ")

    # --- Data Processing ---
    input_filename = "orders_export.csv"
    logging.info(f"Input file: {input_filename}")
    logging.info(f"Processing orders from {start_date.strftime('%d.%m.%Y')} to {end_date.strftime('%d.%m.%Y')}")

    source_df = load_and_validate_csv(input_filename)
    filtered_df = filter_by_date_range(source_df, start_date, end_date)

    final_columns = [
        'Name', 'Fulfilled at', 'Fulfillment Status', 'Financial Status',
        'Created at', 'Lineitem quantity', 'Lineitem name', 'Lineitem sku',
        'Lineitem fulfillment status'
    ]
    existing_columns = [col for col in final_columns if col in filtered_df.columns]
    report_df = filtered_df[existing_columns]

    # --- Report Generation ---
    if not report_df.empty:
        sheets_data = prepare_report_sheets(report_df, cost_first_sku, cost_next_sku, cost_per_piece)
        prompt_message = "Enter the desired name for the output Excel file (e.g., report.xlsx).\nPress Enter to use a default name: "
        output_filename_from_user = input(prompt_message)

        if output_filename_from_user:
            output_filename = output_filename_from_user
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
        else:
            current_date = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"processed_orders_{current_date}.xlsx"
            logging.info(f"No filename provided. Using default: {output_filename}")
        create_excel_report(sheets_data, output_filename)
    else:
        logging.info("Script finished. No orders to process into an Excel file.")

if __name__ == "__main__":
    main()
