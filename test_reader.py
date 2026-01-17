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
    test_importer()