import pandas as pd
import pytest
from shopify_order_processor import (
    calculate_costs,
    transform_cost_df_for_reporting,
    create_invoice_summary
)

from pandas.testing import assert_frame_equal

# --- Constants for Tariffs ---
COST_FIRST_SKU = 1.50
COST_NEXT_SKU = 0.75
COST_PER_PIECE = 0.25


@pytest.fixture
def sample_orders_df():
    """Pytest fixture to create a sample DataFrame for testing."""
    data = {
        'Name': ['#1001', '#1001', '#1002', '#1003', '#1003', '#1003', '#1004'],
        'Lineitem quantity': [1, 2, 3, 1, 1, 1, 1],
        'Lineitem name': ['T-Shirt', 'Mug', 'T-Shirt', 'Hoodie', 'Sticker', 'Package protection', 'Single-Item'],
        'Lineitem sku': ['SKU-TS', 'SKU-MUG', 'SKU-TS', 'SKU-HOOD', 'SKU-STICK', 'INS-01', 'SKU-SINGLE'],
    }
    df = pd.DataFrame(data)
    # Add columns that are expected by the original script but not relevant to the test logic
    df['Fulfilled at'] = pd.NaT
    df['Total'] = 0.0
    return df


def test_calculate_costs_logic(sample_orders_df):
    """
    Tests the new line-item based cost calculation logic.
    This test WILL FAIL with the original code because it does not produce
    the new 'Piece Cost', 'SKU Cost', and 'Line Total Cost' columns.
    """
    # This is the DataFrame we expect after the new logic is implemented
    expected_data = {
        'Piece Cost': [0.25, 0.50, 0.75, 0.25, 0.25, 0.0, 0.25],  # qty * 0.25
        'SKU Cost': [1.50, 0.75, 1.50, 1.50, 0.75, 0.0, 1.50],  # first, next, first, first, next, ignored, first
        'Line Total Cost': [1.75, 1.25, 2.25, 1.75, 1.00, 0.0, 1.75],  # piece + sku
    }
    expected_df = sample_orders_df.copy()
    for col, values in expected_data.items():
        expected_df[col] = values

    # This call to the original `calculate_costs` will fail because it returns
    # different columns and doesn't perform the line-item calculation.
    actual_df = calculate_costs(sample_orders_df, COST_FIRST_SKU, COST_NEXT_SKU, COST_PER_PIECE)

    # We only assert on the columns that are part of the new logic.
    # The test will fail on a KeyError because these columns don't exist in `actual_df`.
    assert_frame_equal(actual_df[expected_df.columns], expected_df)


def test_transform_for_single_item_order():
    """
    Tests that a single-item order is not overwritten by the TOTAL row.
    This test reproduces the bug reported by the user and WILL FAIL with
    the original `transform_cost_df_for_reporting` function.
    """
    # This simulates the data *after* it has been processed by the (new) calculate_costs
    single_order_data = {
        'Name': ['#1004'],
        'Lineitem name': ['Single-Item'],
        'Lineitem quantity': [1],
        'Piece Cost': [0.25],
        'SKU Cost': [1.50],
        'Line Total Cost': [1.75]
    }
    input_df = pd.DataFrame(single_order_data)

    # The original function expects different columns and has different logic,
    # so this call will fail, likely with a KeyError or incorrect output.
    transformed_df = transform_cost_df_for_reporting(input_df)

    # 1. The output should have exactly 2 rows: the item and the total.
    assert len(transformed_df) == 2, "A single-item order should result in a 2-row output (item + TOTAL)"

    # 2. The first row should be the original item, not the total.
    assert transformed_df.iloc[0]['Lineitem name'] == 'Single-Item', "The first row should be the item data"

    # 3. The second row should be the TOTAL row with the correct total.
    assert transformed_df.iloc[1]['Lineitem name'] == 'TOTAL', "The second row should be the TOTAL summary"
    assert transformed_df.iloc[1]['Line Total Cost'] == 1.75, "The TOTAL row should have the correct sum"
