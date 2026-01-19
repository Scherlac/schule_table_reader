#!/usr/bin/env python3
"""
Test script for the ExcelImporter class.
"""

import re
import glob
import os
import pathlib
import pandas as pd
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
    Tests the ExcelImporter class with all Excel files in the current directory.
    Processes each file, places outputs in 'result' subfolder, and creates a summary.xlsx
    with concatenated evaluation data.
    """
    # Find all Excel files in current directory
    excel_files = glob.glob('*.xlsx')
    excel_files = [f for f in excel_files if not f.startswith('~$')]  # Exclude temp files
    
    if not excel_files:
        print("No Excel files found in current directory.")
        return
    
    print(f"Found {len(excel_files)} Excel files: {excel_files}")
    
    # Create result directory if it doesn't exist
    result_dir = 'result'
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)
        print(f"Created result directory: {result_dir}")
    
    # List to collect evaluation DataFrames
    evaluation_dfs = []
    
    # Process each Excel file
    for file_path in excel_files:
        print(f"\nProcessing: {file_path}")
        
        try:
            # Create importer instance
            importer = ExcelImporter(file_path)
            
            if importer.df is None:
                print(f"Failed to load the Excel file: {file_path}")
                continue
            
            # Dump the report
            importer.dump_report()
            
            # Update Excel with statistics (output to result folder)
            base_name = os.path.splitext(file_path)[0]
            output_file = os.path.join(result_dir, f'{base_name}_processed.xlsx')
            importer.update_excel_with_statistics(output_file)
            
            # Get evaluation data
            df_eval = importer.evaluate()
            if df_eval is not None and not df_eval.empty:
                # Add source file column
                df_eval['source_file'] = file_path
                evaluation_dfs.append(df_eval)
                print(f"Collected evaluation data with {len(df_eval.columns)} columns")
            else:
                print("No evaluation data collected")
                
        except Exception as e:
            print(f"Error processing {file_path}: {e}")
            continue
    
    # Concatenate all evaluation DataFrames
    if evaluation_dfs:
        print(f"\nConcatenating {len(evaluation_dfs)} evaluation DataFrames...")
        summary_df = pd.concat(evaluation_dfs, ignore_index=False)
        
        # Save summary to Excel
        summary_path = os.path.join(result_dir, 'summary.xlsx')
        summary_df.to_excel(summary_path)
        
        print(f"Summary saved to: {summary_path}")
        print(f"Summary shape: {summary_df.shape}")
        print("Summary columns:", list(summary_df.columns))
        print("\nFirst few rows of summary:")
        print(summary_df.head())
    else:
        print("No evaluation data collected from any files.")

if __name__ == "__main__":
    test_parse_item_string()
    test_excel_importer()