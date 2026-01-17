#!/usr/bin/env python3
"""
Test script for the ExcelImporter class.
"""

import re
from reader import ExcelImporter, RecordParser

def test_regex_findall():
    """
    Tests the regex findall for extracting item numbers and modifiers.
    """
    item_pattern = r'(?P<num>\d+)(?P<mod>[a-zA-Z]*)'
    
    test_cases = [
        ("22", [("22", "")]),
        ("27n", [("27", "n")]),
        ("Auditív-aktív 22 27n 38 41", [("22", ""), ("27", "n"), ("38", ""), ("41", "")]),
        ("1a 2b", [("1", "a"), ("2", "b")]),
        ("", []),
        ("no digits", []),
    ]
    
    for s, expected in test_cases:
        matches = re.findall(item_pattern, s)
        assert matches == expected, f"For '{s}': expected {expected}, got {matches}"
        print(f"✓ findall test passed for: '{s}' -> {matches}")

def test_regex_sub():
    """
    Tests the regex sub for removing item patterns to extract name.
    """
    item_pattern = r'\d+[a-zA-Z]*'
    
    test_cases = [
        ("Auditív-aktív 22 27n 38 41", "Auditív-aktív"),
        ("Test 1a 2b", "Test"),
        ("Name 10 20", "Name"),
        ("", ""),
        ("Just Name", "Just Name"),
        ("1 2 3", ""),
        ("1a,2b;3c", ",;"),
    ]
    
    for s, expected in test_cases:
        name = re.sub(item_pattern, '', s).strip()
        assert name == expected, f"For '{s}': expected '{expected}', got '{name}'"
        print(f"✓ sub test passed for: '{s}' -> '{name}'")

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
        ("", "", [], []),
        ("Just Name", "Just Name", [], []),
        ("1 2 3", "", [1, 2, 3], ['', '', '']),
        ("1a,2b;3c", ",;", [1, 2, 3], ['a', 'b', 'c']),  # New behavior for ; case
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

    # Get and print child name
    child_name = importer.get_child_name()
    print(f"Child Name: {child_name}")

    # Get and print parsed sections
    parsed_sections = importer.get_parsed_sections()
    print("\nRecord Counts per Section:")
    for section_name, records in parsed_sections.items():
        print(f"{section_name}: {len(records)} records")

    print("\nScore Statistics per Record:")
    for section_name, records in parsed_sections.items():
        print(f"\n{section_name}:")
        for i, record in enumerate(records):
            scores = record['scores']
            if scores:
                count = len(scores)
                total = sum(scores)
                mean = total / count if count > 0 else 0
                print(f"  Record {i+1} ({record['name']}): Count={count}, Sum={total}, Mean={mean:.2f}")
            else:
                print(f"  Record {i+1} ({record['name']}): No scores")

    # Optionally, print full data summary
    full_data = importer.get_full_data()
    print(f"\nFull Data Shape: {full_data.shape}")
    print("First 5 rows of full data:")
    print(full_data.head().to_string(index=False))

if __name__ == "__main__":
    test_regex_findall()
    test_regex_sub()
    test_parse_item_string()
    test_excel_importer()