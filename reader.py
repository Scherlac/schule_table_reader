
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re
import json
import enum
import numpy as np
import math
import shutil
import openpyxl
import os
import pathlib
from typing import List, Dict, Tuple, Optional, Any, Annotated
from pydantic import BaseModel, Field
from models import (
    REPORT_EVAL_ENUM, REPORT_EVAL_TEXT, report_eval_to_enum, report_eval_to_text,
    SectionDetails, RecordDetails, Record, SectionConfig, ClassResult, SectionResult
)


class RecordParser:
    """
    Parser class for extracting structured data from individual Excel rows and sections.
    """
    ECIX : int = 3  # Expected column index for record content (0-indexed, so 3 = column D)

    def __init__(self) -> None:
        pass

    def parse_records(
            self, 
            df: pd.DataFrame, 
            section_name: str, 
            mode: str, 
            start_row_offset: int, 
            full_df: Optional[pd.DataFrame] = None, 
            expected_col_index: int = 3) -> List[Record]:
        """
        Parses the DataFrame into a list of structured records, handling single or multi-row records.

        Parameters:
        df (pd.DataFrame): The section DataFrame.
        mode (str): 'single' for single-row records, 'multi' for multi-row records.
        section_name (str): the name of the section
        start_row_offset (int): The offset to add to row indices to get absolute positions
        full_df (pd.DataFrame): The full DataFrame for RecordDetails
        expected_col_index (int): The expected column index for record content
        Returns:
        list: List of Record objects, each with 'name', 'items', 'modifiers', 'scores'.
        """
        records = []
        pending = None
        subsections = None
        self.ECIX = expected_col_index
        for idx, row in df.iterrows():
            section_relative_idx = idx - start_row_offset  # Section-relative row index
            record = self.parse_record(row, section_relative_idx, start_row_offset, full_df)
            if record:
                record.subsection = subsections
                if mode == 'single':
                    if record.name and record.scores:
                        records.append(record)
                    elif record.name and not record.scores:
                        # subsection header only
                        if record.name != section_name:
                            subsections = record.name
                elif mode == 'multi':
                    if record.name and not record.scores:
                        if pending:
                            # the pending is subsection header only
                            if pending.name != section_name:
                                subsections = pending.name
                        pending = record
                    elif record.scores and pending:
                        # merge pending with record
                        record.name = pending.name
                        records.append(record)
                        pending = None
                    else:
                        pending = None
        return records

    def parse_record(self, row, section_relative_idx: int, section_start_row: int, full_df: Optional[pd.DataFrame] = None) -> Optional[Record]:
        """
        Parses a single row into a structured record.

        Parameters:
        row: The DataFrame row.
        section_relative_idx: The section-relative row index
        section_start_row: The starting row of the section in the full DataFrame
        full_df: The full DataFrame for RecordDetails

        Returns:
        Record: Record object with 'label', 'name', 'items', 'modifiers', 'scores' or None if no label.
        """
        if len(row) > self.ECIX and pd.notna(row[self.ECIX]):
            label = row[self.ECIX]
            name, items, modifiers = self._parse_item_string(str(label))
            scores = [row[i] for i in range(self.ECIX + 1, len(row)) if pd.notna(row[i])]
            details = None
            if full_df is not None:
                details = RecordDetails(
                    source=full_df, 
                    location=(section_relative_idx, self.ECIX)
                )
            return Record(
                label=str(label),
                name=name,
                items=items,
                modifiers=modifiers,
                scores=scores,
                details=details
            )
        return None

    def _parse_item_string(self, s: str) -> Tuple[str, List[int], List[str]]:
        """
        Parses a string to extract name, item numbers, and their modifiers using a single regex pattern.

        Parameters:
        s (str): The string containing the name and item numbers.

        Returns:
        tuple: (name str, list of ints for items, list of strs for modifiers)
        """
        items_pattern = r'''(?x)
            (?P<num>\d+)    # Match item number
            (?P<mod>[a-zA-Z]*) # Match optional modifier
            [ ,;]?          # Optional separator (space, comma, or semicolon)
        '''
        pattern = r'''(?x)
            ^
            (?P<name>.*?)
            [ \s:]*
            (?P<items>
                (
                    (?P<num>\d+)    # Match item number
                    (?P<mod>[a-zA-Z]*) # Match optional modifier
                    [ ,;]*          # Optional separator (space, comma, or semicolon)
                )*
            )
            $
        '''
        matches = re.match(pattern, s)
        if matches:
            name_part = matches.group('name').strip()
            items_part = matches.group('items')
            matches = re.findall(items_pattern, items_part)
            items = [int(match[0]) for match in matches]
            modifiers = [match[1] for match in matches]
            return name_part, items, modifiers
        return s.strip(), [], []


class SectionParser:
    """
    Parser class for extracting structured data from Excel sections.
    """
    ECIX : int = 3  # Expected column index for record content (0-indexed, so 3 = column D)

    def __init__(self, all_sections: List[str], config: SectionConfig) -> None:
        self.record_parser = RecordParser()
        self.all_sections = all_sections
        self.config = config

    def _validate_section_records(self, section_name: str, records: List[Record], config: SectionConfig) -> None:
        """
        Validates the parsed records against the expected configuration.
        More flexible validation that allows for missing or extra records.

        Parameters:
        section_name (str): The name of the section.
        records (List[Record]): The parsed records.
        config (SectionConfig): The configuration for validation.

        Raises:
        ValueError: If validation fails critically.
        """
        if not records:
            # Allow missing sections - just warn and skip
            print(f"Warning: Section {section_name} has no records - skipping evaluation for this section")
            return

        if config.expected_records is not None:
            if len(records) != config.expected_records:
                print(f"Warning: Section {section_name}: expected {config.expected_records} records, got {len(records)} - processing available records")

        if config.expected_questions is not None:
            if len(config.expected_questions) != len(records):
                print(f"Warning: Section {section_name}: expected_questions length {len(config.expected_questions)} != records {len(records)} - using available data")
                # Use min length to avoid index errors
                min_length = min(len(config.expected_questions), len(records))
                config.expected_questions = config.expected_questions[:min_length]
                records = records[:min_length]

            for i, (record, exp) in enumerate(zip(records, config.expected_questions)):
                if len(record.scores) != exp:
                    print(f"Warning: Section {section_name}: record {i+1} expected {exp} questions, got {len(record.scores)} - using available scores")

    def parse_section(
            self, 
            df: pd.DataFrame, 
            section_name: str) -> Optional[SectionResult]:
        """
        Parses the section DataFrame into a structured result with details and records.

        Parameters:
        df (pd.DataFrame): The full DataFrame.
        section_name (str): The name of the section to parse.

        Returns:
        Optional[SectionResult]: Object containing section details and list of records, or None if section not found.
        """
        # Find the exact start and end rows
        start_row = None
        end_row = len(df)
        markers = [s for s in self.all_sections if s != section_name]
        for idx in range(self.config.approx_start, len(df)):
            row = df.iloc[idx]
            for cix, cell in enumerate(row):
                if pd.notna(cell):
                    cell_str = str(cell)
                    if start_row is None and section_name in cell_str:
                        self.ECIX = cix
                        start_row = idx
                    elif start_row is not None and any(marker in cell_str for marker in markers):
                        end_row = idx
                        break
            if end_row != len(df):
                break

        if start_row is None:
            return None

        section_df = df.iloc[start_row:end_row]

        mode = 'multi' if self.config.multi_row else 'single'
        records = self.record_parser.parse_records(section_df, section_name, mode, start_row, df, self.ECIX)
        self._validate_section_records(section_name, records, self.config)
        return SectionResult(
            details=SectionDetails(selection_df=section_df, start_row=start_row, end_row=end_row),
            records=records
        )


class ExcelImporter:
    """
    Importer class to capture and structure the content of the Excel file.
    """
    RCIX : int = 24  # Starting column index for record content (0-indexed, so 24 = column Y)
    ECIX : int = 3   # Expected column index for record content (0-indexed, so 3 = column D)

    def __init__(self, file_path: str) -> None:
        """
        Initializes the importer by reading the Excel file.

        Parameters:
        file_path (str): The path to the Excel file.
        """
        # try:
        if True:
            self.file_path = file_path
            self.df = pd.read_excel(file_path, header=None)
            self.sections : Dict[str, pd.DataFrame] = {}
            self.parsed_sections : Dict[str, SectionResult] = {}
            self.child_name : Optional[str] = None

            self.sections_config : Dict[str, SectionConfig] = {}

            

            # Load section configuration from JSON
            self._load_config('sections_config.json')

            self._extract_data()
        # except Exception as e:
        #     print(f"An error occurred while importing the Excel file: {e}")
        #     self.df = None

    def _load_config(self, config_path: str) -> None:
        """
        Loads the section configuration from a JSON file.

        Parameters:
        config_path (str): The path to the JSON configuration file.
        """
        with open(config_path, 'r', encoding='utf-8') as f:
            raw_config = json.load(f)
            self.sections_config = {k: SectionConfig.model_validate(v) for k, v in raw_config.items()}

    def _extract_data(self) -> None:
        """
        Extracts the child name and sections from the DataFrame.
        """
        if self.df is None:
            return

        # Extract child name (assuming in row 1, column 3 - 0-indexed row 1, col 3)
        if len(self.df) > 1 and len(self.df.columns) > 3:
            self.child_name = self.df.iloc[1, self.ECIX] if pd.notna(self.df.iloc[1, self.ECIX]) else None

        all_sections = list(self.sections_config.keys())

        # Parse sections using configuration
        for section_name, config in self.sections_config.items():
            parser = SectionParser(all_sections, config)
            section = parser.parse_section(self.df, section_name)
            if section:
                self.parsed_sections[section_name] = section

    def get_child_name(self) -> Optional[str]:
        """
        Returns the child name.

        Returns:
        str: The child name or None if not found.
        """
        return self.child_name

    def get_sections(self) -> Dict[str, pd.DataFrame]:
        """
        Returns the sections data.

        Returns:
        dict: Dictionary with section names as keys and DataFrames as values.
        """
        return self.sections

    def get_parsed_sections(self) -> Dict[str, SectionResult]:
        """
        Returns the parsed sections data.

        Returns:
        dict: Dictionary with section names as keys and parsed data as values.
        """
        return self.parsed_sections

    def dump_report(self) -> None:
        """
        Dumps a comprehensive report of the imported Excel data.
        """
        if self.df is None:
            print("No data loaded.")
            return

        # Get child name
        child_name = self.get_child_name()
        print(f"Child Name: {child_name}")

        # Get parsed sections
        parsed_sections = self.get_parsed_sections()
        print("\nRecord Counts per Section:")
        for section_name, section_result in parsed_sections.items():
            print(f"{section_name}: {len(section_result.records)} records")

        print("\nScore Statistics per Record:")
        for section_name, section_result in parsed_sections.items():
            print(f"\n{section_name}:")
            for i, record in enumerate(section_result.records):
                scores = record.scores
                subsection = record.subsection
                name = record.name
                display_name = f"{subsection}/{name}" if subsection else name
                if scores:
                    count = len(scores)
                    total = np.sum(scores)
                    mean = total / count if count > 0 else 0
                    print(f"  Record {i+1} ({display_name}): Count={count}, Sum={total}, Mean={mean:.2f}")
                else:
                    print(f"  Record {i+1} ({display_name}): No scores")

    def update_excel_with_record_statistics(self, section_name : str, section_result: SectionResult) -> None:
            
            section_start_row = section_result.details.start_row
            selection_config : SectionConfig  = self.sections_config[section_name]
            class_eval_flags = selection_config.class_eval_flags
            record_eval_flags = selection_config.record_eval_flags
            expected_questions = self.sections_config[section_name].expected_questions


            # Add column headers at section start row
            self.df.iloc[section_start_row, self.RCIX + 0] = "Sum"
            self.df.iloc[section_start_row, self.RCIX + 1] = "Mean"
            self.df.iloc[section_start_row, self.RCIX + 2] = "Normed Sum"
            
            # Process each record in the section
            for i, record in enumerate(section_result.records):
                if record.scores and record.details:
                    scores = np.array(record.scores).astype(float)
                    entered_questions = len(scores)
                    expected_question = expected_questions[i] 
                    print(f"Processing record '{record.name}' with scores: {scores}, expected questions: {expected_question}, entered questions: {entered_questions}")
                    total = scores.sum()
                    mean = scores.mean() 

                    normed_sum = max(expected_question, entered_questions) * mean
                    # convert to int
                    normed_sum = int(round(normed_sum))

                    record.eval_results = {
                        REPORT_EVAL_TEXT.SUM: total,
                        REPORT_EVAL_TEXT.MEAN: mean,
                        REPORT_EVAL_TEXT.NORMED_SUM: normed_sum
                    }
                    
                    # Get the section-relative row and add section_start_row for absolute position
                    section_relative_row, col_idx = record.details.location
                    row_idx = section_relative_row + section_result.details.start_row
                    
                    # Column U (20) for Sum, Column V (21) for Mean
                    self.df.iloc[row_idx, self.RCIX + 0] = total
                    self.df.iloc[row_idx, self.RCIX + 1] = mean
                    self.df.iloc[row_idx, self.RCIX + 2] = normed_sum

    def update_excel_with_subsection_statistics(self, section_name : str, section_result: SectionResult) -> None:
        section_start_row = section_result.details.start_row

        # Add column headers at section start row
        self.df.iloc[section_start_row, self.RCIX + 3] = "Grade limit"
        self.df.iloc[section_start_row, self.RCIX + 4] = "Class evaluation"
        self.df.iloc[section_start_row, self.RCIX + 5] = "Record evaluation"

        classification = self.sections_config[section_name].classification
        question_id = self.sections_config[section_name].question_id
        class_marker = self.sections_config[section_name].class_marker
        if not classification or not class_marker:
            return
        
        # Process each class
        class_results : Dict[str, ClassResult] = section_result.class_results
        for class_id in set(classification):
            marker = class_marker[class_id - 1]  # assuming class_id starts from 1
            records_in_class = [(i, idx, section_result.records[i]) for i, (c, idx) in enumerate(zip(classification, question_id)) if c == class_id]

            try:
                means = np.array([r.eval_results[REPORT_EVAL_TEXT.MEAN] for i, idx, r in records_in_class if r.eval_results and REPORT_EVAL_TEXT.MEAN in r.eval_results]) 
            except:
                means = np.array([])

            mean_of_means = means.mean() 
            std_of_means = means.std()
            max_of_means = means.max()

            high_grade_limit = (1.0 - 0.001) * (max_of_means - std_of_means) # 0.001 error margin

            selected_records = []

            for i, idx, record in records_in_class:

                section_relative_row, col_idx = record.details.location
                row_idx = section_relative_row + section_result.details.start_row

                if record.eval_results and REPORT_EVAL_TEXT.MEAN in record.eval_results:
                    is_high = record.eval_results[REPORT_EVAL_TEXT.MEAN] >= high_grade_limit
                    if is_high:
                        selected_records.append(idx)

                    # add to excel
                    self.df.iloc[row_idx, self.RCIX + 5] = "High" if is_high else "Other"

            class_evaluation = ', '.join(map(str, selected_records))

            i, idx, first_record = records_in_class[0]
            section_relative_row, col_idx = first_record.details.location
            row_idx = section_relative_row + section_result.details.start_row

            self.df.iloc[row_idx, self.RCIX + 3] = round(high_grade_limit, 2)
            self.df.iloc[row_idx, self.RCIX + 4] = f"{marker}: {class_evaluation}"
            class_results[marker] = ClassResult(
                class_id=class_id,
                marker=marker,
                records=[r for i, idx, r in records_in_class],
                eval_results={
                    REPORT_EVAL_TEXT.GRADE_LIMIT : high_grade_limit,
                    REPORT_EVAL_TEXT.HIGH_GRADES : class_evaluation,
                    REPORT_EVAL_TEXT.MEAN : mean_of_means,
                    REPORT_EVAL_TEXT.STD : std_of_means,
                    REPORT_EVAL_TEXT.MAX : max_of_means,
                })



    def update_excel_with_statistics(self, output_file: str) -> None:
        """
        Calculates sum and mean for all records and updates the Excel file with results in columns S and T.
        Adds column headers at section start rows.
        
        Parameters:
        output_file (str): The path to save the updated Excel file. Can include a folder path.
        """
        if self.df is None:
            print("No data loaded.")
            return

        # Ensure the output directory exists
        self.output_file = pathlib.Path(output_file)
        self.output_dir = self.output_file.parent.absolute()
        if not self.output_dir.exists():
            print(f"Creating output directory: {str(self.output_dir)}")
            self.output_dir.mkdir(parents=True, exist_ok=True)

        # Copy the original Excel file to the output location first
        shutil.copy2(self.file_path, output_file)
        print(f"Copied original Excel file to: {output_file}")

        # Ensure the DataFrame has at least 20 columns (0-19, where 18 is column S, 19 is column T)
        while len(self.df.columns) <= self.RCIX + 7:
            self.df[len(self.df.columns)] = None

        # Get all parsed sections
        parsed_sections = self.get_parsed_sections()
        
        # Process each section
        for section_name, section_result in parsed_sections.items():

            print(f"Updating statistics for section: {section_name}")

            self.update_excel_with_record_statistics(section_name, section_result)

            self.update_excel_with_subsection_statistics(section_name, section_result)
            

        # Update the copied Excel file with the modified data
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        
        # Write the DataFrame data to the worksheet, preserving formatting
        for row_idx in range(len(self.df)):
            for col_idx in range(len(self.df.columns)):
                value = self.df.iloc[row_idx, col_idx]
                # Convert numpy types to native Python types for Excel
                if pd.isna(value):
                    value = None
                elif hasattr(value, 'item'):  # numpy scalar
                    value = value.item()
                ws.cell(row=row_idx + 1, column=col_idx + 1, value=value)
        
        wb.save(output_file)
        print(f"Updated Excel saved to: {output_file}")

    def evaluate(self) -> pd.DataFrame:
        """
        Placeholder for future evaluation logic.
        """
        result = {}

        for section_name, section_result in self.parsed_sections.items():
            selection_config : SectionConfig  = self.sections_config[section_name]
            class_eval = [s.strip() for s in (selection_config.class_eval or '').split(',')]
            record_eval = [s.strip() for s in (selection_config.record_eval or '').split(',')]
            print(f"Evaluating section: {section_name}")

            #iterate over the classes and append evaluation results
            for marker in set(selection_config.class_marker or []):
                if marker not in section_result.class_results:
                    continue
                class_result = section_result.class_results[marker]

                # iterate over the flags and append evaluation results
                for label in class_eval:
                    if label not in report_eval_to_enum:
                        continue
                    value = label
                    if value in class_result.eval_results:
                        result[f"{marker}: {label}"] = class_result.eval_results[value]

            # iterate over the records and append evaluation results
            record : Record = None
            for i, record in enumerate(section_result.records):
                if selection_config.classification:
                    class_id = selection_config.classification[i]
                    marker = selection_config.class_marker[class_id - 1]
                else:
                    marker = record.name
                for label in record_eval:
                    if label not in report_eval_to_enum:
                        continue
                    value = REPORT_EVAL_TEXT(label)
                    if value in record.eval_results:
                        result[f"{marker} - {record.name}: {label}"] = record.eval_results[value]

        # If using all scalar values, you must pass an index
        # child_name as index
        if self.child_name is None:
            self.child_name = "Unknown"
        
        return pd.DataFrame(result, index=[self.child_name])



    
