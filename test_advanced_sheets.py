#!/usr/bin/env python3
"""
Test script for Advanced Sheets tools in Google Workspace MCP v9.
Tests all 13 new tools against a test spreadsheet.
"""

import sys
import os

# Add the MCP server directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from server import (
    google_client,
    sheets_batch_update,
    rename_sheet,
    format_cells,
    set_column_width,
    set_row_height,
    freeze_rows_columns,
    merge_cells,
    unmerge_cells,
    add_filter,
    data_validation,
    conditional_formatting,
    named_range,
    auto_resize,
    add_sheet_tab,
    delete_sheet_tab,
    get_sheet_metadata
)
import asyncio

# Test spreadsheet ID
SPREADSHEET_ID = "1qijTLBSVAhmDSzZCLUVQWdpN7SLBBWO-bYVv0aFwO0w"
TEST_SHEET_NAME = "_TEST_Advanced_Features_v9"
TEST_SHEET_ID = None  # Will be set after creating the test sheet


async def create_test_sheet():
    """Create a test sheet and return its ID."""
    global TEST_SHEET_ID
    result = await add_sheet_tab({
        'spreadsheet_id': SPREADSHEET_ID,
        'title': TEST_SHEET_NAME
    })
    print(f"✓ Created test sheet: {result[0].text}")

    # Get the sheet ID from metadata
    metadata = await get_sheet_metadata({'spreadsheet_id': SPREADSHEET_ID})
    for line in metadata[0].text.split('\n'):
        if TEST_SHEET_NAME in line:
            # Extract sheet ID from "- _TEST_Advanced_Features_v9 (ID: 123456, ..."
            import re
            match = re.search(r'ID: (\d+)', line)
            if match:
                TEST_SHEET_ID = int(match.group(1))
                print(f"  Sheet ID: {TEST_SHEET_ID}")
                break

    return TEST_SHEET_ID


async def write_test_data():
    """Write some test data to the sheet."""
    # Use the sheets service directly to write test data
    values = [
        ["Name", "Status", "Priority", "Score", "Notes"],
        ["Task 1", "In Progress", "High", "85", "First task"],
        ["Task 2", "Done", "Medium", "92", "Second task"],
        ["Task 3", "Pending", "Low", "78", "Third task"],
        ["Task 4", "In Progress", "High", "65", "Fourth task"],
        ["Task 5", "Done", "Medium", "88", "Fifth task"],
    ]

    body = {"values": values}
    google_client.sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{TEST_SHEET_NAME}!A1:E6",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    print("✓ Wrote test data to sheet")


async def test_format_cells():
    """Test formatting cells - make headers bold with background color."""
    result = await format_cells({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'range': 'A1:E1',
        'bold': True,
        'background_color': {'red': 0.2, 'green': 0.4, 'blue': 0.8},
        'text_color': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
        'horizontal_alignment': 'CENTER'
    })
    print(f"✓ Format cells: {result[0].text}")


async def test_set_column_width():
    """Test setting column widths."""
    result = await set_column_width({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_column': 0,  # Column A
        'end_column': 1,    # Just column A
        'width': 150
    })
    print(f"✓ Set column width: {result[0].text}")


async def test_set_row_height():
    """Test setting row heights."""
    result = await set_row_height({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 0,  # Row 1 (header)
        'end_row': 1,
        'height': 35
    })
    print(f"✓ Set row height: {result[0].text}")


async def test_freeze_rows_columns():
    """Test freezing rows and columns."""
    result = await freeze_rows_columns({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'frozen_rows': 1,
        'frozen_columns': 1
    })
    print(f"✓ Freeze rows/columns: {result[0].text}")


async def test_data_validation():
    """Test adding data validation (dropdown) to Status column."""
    result = await data_validation({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 1,     # Skip header
        'end_row': 10,      # Rows 2-10
        'start_column': 1,  # Column B (Status)
        'end_column': 2,
        'validation_type': 'ONE_OF_LIST',
        'values': ['Pending', 'In Progress', 'Done', 'Blocked']
    })
    print(f"✓ Data validation: {result[0].text}")


async def test_conditional_formatting():
    """Test adding conditional formatting to Status column."""
    # Green for "Done"
    result1 = await conditional_formatting({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 1,
        'end_row': 10,
        'start_column': 1,  # Column B
        'end_column': 2,
        'rule_type': 'TEXT_EQ',
        'values': ['Done'],
        'background_color': {'red': 0.7, 'green': 0.9, 'blue': 0.7}
    })
    print(f"✓ Conditional formatting (Done=green): {result1[0].text}")

    # Red for "Blocked"
    result2 = await conditional_formatting({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 1,
        'end_row': 10,
        'start_column': 1,
        'end_column': 2,
        'rule_type': 'TEXT_EQ',
        'values': ['Blocked'],
        'background_color': {'red': 0.9, 'green': 0.7, 'blue': 0.7}
    })
    print(f"✓ Conditional formatting (Blocked=red): {result2[0].text}")


async def test_add_filter():
    """Test adding a filter to the data range."""
    result = await add_filter({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 0,
        'start_column': 0,
        'end_column': 5
    })
    print(f"✓ Add filter: {result[0].text}")


async def test_merge_cells():
    """Test merging cells (avoiding frozen column A)."""
    result = await merge_cells({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 7,  # Row 8
        'end_row': 8,    # Row 9 (exclusive)
        'start_column': 1,  # Start at B (avoid frozen column A)
        'end_column': 5,
        'merge_type': 'MERGE_ALL'
    })
    print(f"✓ Merge cells: {result[0].text}")


async def test_unmerge_cells():
    """Test unmerging cells (matching the merged range)."""
    result = await unmerge_cells({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'start_row': 7,
        'end_row': 8,
        'start_column': 1,  # Match the merged range
        'end_column': 5
    })
    print(f"✓ Unmerge cells: {result[0].text}")


async def test_named_range():
    """Test creating a named range."""
    result = await named_range({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'name': 'TestDataRange',
        'start_row': 0,
        'end_row': 6,
        'start_column': 0,
        'end_column': 5
    })
    print(f"✓ Named range: {result[0].text}")


async def test_auto_resize():
    """Test auto-resizing columns."""
    result = await auto_resize({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'dimension': 'COLUMNS',
        'start_index': 0,
        'end_index': 5
    })
    print(f"✓ Auto resize: {result[0].text}")


async def test_rename_sheet():
    """Test renaming the sheet."""
    result = await rename_sheet({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'new_title': '_TEST_Advanced_RENAMED'
    })
    print(f"✓ Rename sheet: {result[0].text}")

    # Rename it back
    result = await rename_sheet({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID,
        'new_title': TEST_SHEET_NAME
    })
    print(f"✓ Rename sheet (back): {result[0].text}")


async def test_batch_update():
    """Test batch update with multiple operations."""
    result = await sheets_batch_update({
        'spreadsheet_id': SPREADSHEET_ID,
        'requests': [
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': TEST_SHEET_ID,
                        'dimension': 'COLUMNS',
                        'startIndex': 4,  # Column E
                        'endIndex': 5
                    },
                    'properties': {
                        'pixelSize': 200
                    },
                    'fields': 'pixelSize'
                }
            }
        ]
    })
    print(f"✓ Batch update: {result[0].text}")


async def cleanup_test_sheet():
    """Delete the test sheet."""
    result = await delete_sheet_tab({
        'spreadsheet_id': SPREADSHEET_ID,
        'sheet_id': TEST_SHEET_ID
    })
    print(f"✓ Cleanup: {result[0].text}")


async def main():
    print("=" * 60)
    print("Testing Advanced Sheets Tools (v9)")
    print("=" * 60)
    print(f"Spreadsheet ID: {SPREADSHEET_ID}")
    print()

    # Authenticate
    if not google_client.authenticate():
        print("❌ Failed to authenticate")
        return 1
    print("✓ Authenticated with Google APIs")
    print()

    try:
        # Create test sheet
        print("--- Creating Test Sheet ---")
        await create_test_sheet()
        await write_test_data()
        print()

        # Run tests
        print("--- Testing Format Operations ---")
        await test_format_cells()
        await test_set_column_width()
        await test_set_row_height()
        print()

        print("--- Testing Freeze ---")
        await test_freeze_rows_columns()
        print()

        print("--- Testing Data Validation ---")
        await test_data_validation()
        print()

        print("--- Testing Conditional Formatting ---")
        await test_conditional_formatting()
        print()

        print("--- Testing Filter ---")
        await test_add_filter()
        print()

        print("--- Testing Merge/Unmerge ---")
        await test_merge_cells()
        await test_unmerge_cells()
        print()

        print("--- Testing Named Range ---")
        await test_named_range()
        print()

        print("--- Testing Auto Resize ---")
        await test_auto_resize()
        print()

        print("--- Testing Rename Sheet ---")
        await test_rename_sheet()
        print()

        print("--- Testing Batch Update ---")
        await test_batch_update()
        print()

        print("=" * 60)
        print("All tests passed! ✓")
        print("=" * 60)

        # Ask if user wants to keep or delete the test sheet
        print()
        print(f"Test sheet '{TEST_SHEET_NAME}' (ID: {TEST_SHEET_ID}) was created.")
        response = input("Delete test sheet? (y/n): ").strip().lower()
        if response == 'y':
            print()
            print("--- Cleanup ---")
            await cleanup_test_sheet()
        else:
            print(f"Test sheet kept. View it at:")
            print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={TEST_SHEET_ID}")

        return 0

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()

        # Try to cleanup on failure
        if TEST_SHEET_ID:
            try:
                print("\nAttempting cleanup...")
                await cleanup_test_sheet()
            except:
                pass

        return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
