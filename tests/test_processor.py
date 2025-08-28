import pandas as pd
import pytest
import shutil
from pathlib import Path

import os
from shopify_order_processor import main as process_orders_main


def test_end_to_end_run(tmp_path, mocker):
    """
    Tests the full script execution from start to finish.
    This test simulates user input, runs the main() function,
    and inspects the output Excel file to verify correctness, specifically
    checking for the bug where single-item orders are overwritten.
    """
    # 1. Setup the test environment
    # Source path for test data
    test_data_path = Path(__file__).parent / "data" / "sample_orders.csv"
    # The script expects the input file in the CWD
    script_run_dir = tmp_path
    shutil.copy(test_data_path, script_run_dir / "orders_export.csv")

    # Store original CWD and change to the temporary directory for the test run
    original_cwd = Path.cwd()
    os.chdir(script_run_dir)

    # 2. Mock user input
    # The script will ask for: start date, end date, 3 tariffs, and output filename
    mock_inputs = [
        "01.08.2025",   # Start Date
        "03.08.2025",   # End Date
        "1.50",         # Cost for first SKU
        "0.75",         # Cost for subsequent SKU
        "0.25",         # Cost per piece
        "output.xlsx",  # Output filename
    ]
    mocker.patch('builtins.input', side_effect=mock_inputs)
    # Also mock print to keep test output clean
    mocker.patch('builtins.print')

    # 3. Run the script's main function
    # This will fail with the original code, as it doesn't implement the new logic.
    process_orders_main()

    # 4. Assert on the output
    output_file = Path("output.xlsx")
    assert output_file.exists(), "The output Excel file was not created."

    # Read the relevant sheet from the generated Excel file
    # We expect a failure here until the script is fixed, as the columns will be wrong.
    try:
        df = pd.read_excel(output_file, sheet_name="Cost Calculation")
    except Exception as e:
        # Restore CWD before failing
        os.chdir(original_cwd)
        pytest.fail(f"Could not read the 'Cost Calculation' sheet from the output Excel. Error: {e}")


    # --- Verification for Single-Item Order (#1003) ---
    order_1003 = df[df['Name'] == '#1003'].reset_index(drop=True)
    assert len(order_1003) == 2, "Order #1003 (single item) should have 2 rows (item + TOTAL)"
    assert order_1003.iloc[0]['Lineitem name'] == 'Single-Item'
    assert order_1003.iloc[1]['Lineitem name'] == 'TOTAL'
    # Check the calculated total for the single item order
    # Expected: (3 pieces * 0.25) + (1 first SKU * 1.50) = 0.75 + 1.50 = 2.25
    assert order_1003.iloc[1]['Line Total Cost'] == pytest.approx(2.25)

    # --- Verification for Multi-Item Order (#1001) ---
    order_1001 = df[df['Name'] == '#1001'].reset_index(drop=True)
    assert len(order_1001) == 3, "Order #1001 (multi-item) should have 3 rows (2 items + TOTAL)"
    assert order_1001.iloc[2]['Lineitem name'] == 'TOTAL'
    # Expected: (1*0.25 + 1.50) + (2*0.25 + 0.75) = 1.75 + 1.25 = 3.00
    assert order_1001.iloc[2]['Line Total Cost'] == pytest.approx(3.00)

    # --- Verification for Order with Protection (#1002) ---
    order_1002 = df[df['Name'] == '#1002'].reset_index(drop=True)
    assert len(order_1002) == 3, "Order #1002 (with protection) should have 3 rows (2 items + TOTAL)"
    # Check that the "Package protection" item has 0 cost
    protection_row = order_1002[order_1002['Lineitem name'] == 'Package protection']
    assert not protection_row.empty, "Package protection row should be present"
    assert protection_row.iloc[0]['Line Total Cost'] == 0.0, "Package protection item should have a total cost of 0"
    # Check the total for the order (only the hoodie should be billed)
    # Expected: (1 piece * 0.25) + (1 first SKU * 1.50) = 1.75
    assert order_1002.iloc[2]['Line Total Cost'] == pytest.approx(1.75)

    # Restore CWD
    os.chdir(original_cwd)
