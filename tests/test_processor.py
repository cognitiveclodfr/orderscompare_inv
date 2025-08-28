import pandas as pd
import pytest
import shutil
import os
from pathlib import Path

from shopify_order_processor import main as process_orders_main

def test_end_to_end_with_user_data(tmp_path, mocker):
    """
    Tests the full script execution using the user-provided CSV data to
    reproduce and verify the fix for the order grouping bug.
    """
    # 1. Setup the test environment using user's data
    test_data_path = Path(__file__).parent / "data" / "user_provided_orders.csv"
    script_run_dir = tmp_path
    shutil.copy(test_data_path, script_run_dir / "orders_export.csv")

    original_cwd = Path.cwd()
    os.chdir(script_run_dir)

    # 2. Mock user input
    mock_inputs = [
        "01.06.2025",   # Start Date
        "15.06.2025",   # End Date
        "0.87",         # Cost for first SKU
        "0.31",         # Cost for subsequent SKU
        "0.24",         # Cost per piece
        "output.xlsx",  # Output filename
    ]
    mocker.patch('builtins.input', side_effect=mock_inputs)
    mocker.patch('builtins.print')

    # 3. Run the script's main function
    # This will fail with the original code, reproducing the user's issue.
    process_orders_main()

    # 4. Assert on the output
    output_file = Path("output.xlsx")
    assert output_file.exists(), "The output Excel file was not created."

    try:
        df = pd.read_excel(output_file, sheet_name="Cost Calculation")
    except Exception as e:
        os.chdir(original_cwd)
        pytest.fail(f"Could not read 'Cost Calculation' sheet. Error: {e}")

    # Restore CWD after file operations
    os.chdir(original_cwd)

    # --- Verification for Order #129711 (multi-line, was split) ---
    order_129711 = df[df['Name'] == '#129711']
    # Should be 6 items + 1 TOTAL row
    assert len(order_129711) == 7, "Order #129711 should have 7 rows (6 items + 1 TOTAL)"
    assert order_129711.iloc[-1]['Lineitem name'] == 'TOTAL', "The last row for #129711 must be the TOTAL row"

    # --- Verification for Order #129715 (single-line, was inverted) ---
    order_129715 = df[df['Name'] == '#129715']
    # Should be 1 item + 1 TOTAL row
    assert len(order_129715) == 2, "Order #129715 should have 2 rows (1 item + 1 TOTAL)"
    assert order_129715.iloc[0]['Lineitem name'] != 'TOTAL', "The first row for #129715 must be the item row"
    assert order_129715.iloc[1]['Lineitem name'] == 'TOTAL', "The second row for #129715 must be the TOTAL row"

    # --- Verification for Order #129807 (another single-line) ---
    order_129807 = df[df['Name'] == '#129807']
    assert len(order_129807) == 2, "Order #129807 should have 2 rows (1 item + 1 TOTAL)"
    assert order_129807.iloc[1]['Lineitem name'] == 'TOTAL', "The last row for #129807 must be the TOTAL row"
