#!/usr/bin/env python3
"""
Test script for the ExcelImporter class.
"""

import re
from reader import ExcelImporter, RecordParser


def test_parse_item_string():
    """
    Tests the _parse_item_string method with various inputs.
    """
    parser = RecordParser()
    
    test_cases = [
        # (input_string, expected_name, expected_items, expected_modifiers)
        ("Auditív-aktív 22 27n 38 41", "Auditív-aktív", [22, 27, 38, 41], ['', 'n', '', '']),
        ("Test 1a 2b", "Test", [1, 2], ['a', 'b']),
        ("Name 10 20", "Name", [10, 20], ['', '']),
        ("– affiliáció (M3): az odatartozás szükséglete, főleg egykorúakhoz",
         "– affiliáció (M3): az odatartozás szükséglete, főleg egykorúakhoz", [], []),
        ("", "", [], []),
        ("Just Name", "Just Name", [], []),
        ("1 2 3", "", [1, 2, 3], ['', '', '']),
        ("1a,2b;3c", "", [1, 2, 3], ['a', 'b', 'c']),  # New behavior for ; case
    ]
    
    for s, exp_name, exp_items, exp_modifiers in test_cases:
        name, items, modifiers = parser._parse_item_string(s)
        assert name == exp_name, f"For '{s}': expected name '{exp_name}', got '{name}'"
        assert items == exp_items, f"For '{s}': expected items {exp_items}, got {items}"
        assert modifiers == exp_modifiers, f"For '{s}': expected modifiers {exp_modifiers}, got {modifiers}"
        print(f"✓ Test passed for: '{s}' -> name='{name}', items={items}, modifiers={modifiers}")

def test_excel_importer():
    """
    Tests the ExcelImporter class with the sample Excel file.
    """
    file_path = 'példatáblázat.xlsx'

    # Create importer instance
    importer = ExcelImporter(file_path)

    if importer.df is None:
        print("Failed to load the Excel file.")
        return

    # Dump the report
    importer.dump_report()

if __name__ == "__main__":
    test_parse_item_string()
    test_excel_importer()