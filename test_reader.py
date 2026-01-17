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

    # Get and print sections
    sections = importer.get_sections()
    print("\nSections found:")
    for section_name, section_df in sections.items():
        print(f"\n{section_name} Section:")
        print(section_df.to_string(index=False))

    # Optionally, print full data summary
    full_data = importer.get_full_data()
    print(f"\nFull Data Shape: {full_data.shape}")
    print("First 5 rows of full data:")
    print(full_data.head().to_string(index=False))

if __name__ == "__main__":
    test_importer()