#!/usr/bin/env python3
"""
Test script for the ExcelImporter class.
"""

from reader import ExcelImporter

def test_importer():
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
    print("\nParsed Sections:")
    for section_name, records in parsed_sections.items():
        print(f"\n{section_name} Parsed (first 3 records):")
        for record in records[:3]:
            print(f"  Name: {record['name']}, Items: {record['items']}, Modifiers: {record['modifiers']}, Scores: {record['scores']}")

    # Optionally, print full data summary
    full_data = importer.get_full_data()
    print(f"\nFull Data Shape: {full_data.shape}")
    print("First 5 rows of full data:")
    print(full_data.head().to_string(index=False))

if __name__ == "__main__":
    test_importer()