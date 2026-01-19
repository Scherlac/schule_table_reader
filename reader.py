
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re
import json
import enum
import numpy as np
import math
import shutil
import openpyxl
from typing import List, Dict, Tuple, Optional, Any, Annotated
from pydantic import BaseModel, Field


class REPORT_EVAL_ENUM(enum.Flag):
    HIGH_GRADES = enum.auto()
    HIGH_GRADES_TOTAL = enum.auto()
    MEAN = enum.auto()
    SUM = enum.auto()
    NORMED_SUM = enum.auto()
    COPY = enum.auto()

report_eval_dict = {
    "> max - dev": REPORT_EVAL_ENUM.HIGH_GRADES,
    "> max - dev (total)": REPORT_EVAL_ENUM.HIGH_GRADES_TOTAL,
    "mean": REPORT_EVAL_ENUM.MEAN,
    "sum": REPORT_EVAL_ENUM.SUM,
    "normed sum": REPORT_EVAL_ENUM.NORMED_SUM,
    "copy": REPORT_EVAL_ENUM.COPY
}


class SectionDetails(BaseModel):
    selection_df: Any = Field(description="The DataFrame containing the selected section data")
    start_row: int = Field(description="The starting row index of the section")
    end_row: int = Field(description="The ending row index of the section")


class RecordDetails(BaseModel):
    source: Any = Field(description="The DataFrame containing the source data")
    location: Tuple[int, int] = Field(description="The (section-relative row, column) coordinates")


class Record(BaseModel):
    label: str = Field(description="The original label string from the Excel row")
    name: str = Field(description="The parsed name of the record")
    items: List[int] = Field(description="List of item numbers associated with the record")
    modifiers: List[str] = Field(description="List of modifiers for each item")
    scores: List[Any] = Field(description="List of score values for the record")
    subsection: Optional[str] = Field(None, description="Optional subsection name if the record belongs to a subsection")
    details: Optional[RecordDetails] = Field(None, description="Details about the source and location of the data")
    eval_results: Optional[Dict[str, Any]] = Field(None, description="Evaluation results for the record")



class SectionConfig(BaseModel):
    approx_start: int = Field(description="Approximate starting row for searching the section")
    multi_row: bool = Field(description="Whether records in this section span multiple rows")
    expected_records: Optional[int] = Field(None, description="Expected number of records in the section")
    expected_questions: Optional[List[int]] = Field(None, description="Expected number of questions per record")
    # "classification": [1, 1, 1, 1, 1, 2, 2, 3, 3],
    classification: Optional[List[int]] = Field(None, description="Class identifiers for each record")
    # "class_marker": ["cognitive", "social", "logic"],
    class_marker: Optional[List[str]] = Field(None, description="Class markers for grouping records")
    # "question_id": [11, 12, 21, 22, 31, 41, 42, 51, 52],
    question_id: Optional[List[int]] = Field(None, description="Question IDs for each record")
    # "report_eval": "> max - dev"
    class_eval: Optional[str] = Field(None, description="Evaluation criteria for reporting scores on class level")
    record_eval: Optional[str] = Field(None, description="Evaluation criteria for reporting scores on record level")

    @property
    def class_eval_flags(self) -> REPORT_EVAL_ENUM:
        """
        Parses the class_eval string into a combined REPORT_EVAL_ENUM flag.

        Returns:
        REPORT_EVAL_ENUM: Combined flags based on the class_eval configuration.
        """
        if not self.class_eval:
            return REPORT_EVAL_ENUM(0)
        flags = REPORT_EVAL_ENUM(0)
        parts = [part.strip() for part in self.class_eval.split(',')]
        for part in parts:
            if part in report_eval_dict:
                flags |= report_eval_dict[part]
        return flags
    
    @class_eval_flags.setter
    def class_eval_flags(self, flags: REPORT_EVAL_ENUM) -> None:
        """
        Sets the class_eval string based on the provided REPORT_EVAL_ENUM flags.

        Parameters:
        flags (REPORT_EVAL_ENUM): Combined flags to set the class_eval configuration.
        """
        parts = []
        for key, value in report_eval_dict.items():
            if flags & value:
                parts.append(key)
        self.class_eval = ', '.join(parts)

    @property
    def record_eval_flags(self) -> REPORT_EVAL_ENUM:
        """
        Parses the record_eval string into a combined REPORT_EVAL_ENUM flag.

        Returns:
        REPORT_EVAL_ENUM: Combined flags based on the record_eval configuration.
        """
        if not self.record_eval:
            return REPORT_EVAL_ENUM(0)
        flags = REPORT_EVAL_ENUM(0)
        parts = [part.strip() for part in self.record_eval.split(',')]
        for part in parts:
            if part in report_eval_dict:
                flags |= report_eval_dict[part]
        return flags
    
    @record_eval_flags.setter
    def record_eval_flags(self, flags: REPORT_EVAL_ENUM) -> None:
        """
        Sets the record_eval string based on the provided REPORT_EVAL_ENUM flags.

        Parameters:
        flags (REPORT_EVAL_ENUM): Combined flags to set the record_eval configuration.
        """
        parts = []
        for key, value in report_eval_dict.items():
            if flags & value:
                parts.append(key)
        self.record_eval = ', '.join(parts)



class SectionResult(BaseModel):
    details: SectionDetails = Field(description="Details about the section extraction")
    records: List[Record] = Field(description="List of parsed records from the section")
    # subsections: Optional[Dict[str, List[Record]]] = Field(default=None, description="Optional mapping of subsection names to their records")


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

        Parameters:
        section_name (str): The name of the section.
        records (List[Record]): The parsed records.
        config (SectionConfig): The configuration for validation.

        Raises:
        ValueError: If validation fails.
        """
        if config.expected_records is not None and len(records) != config.expected_records:
            raise ValueError(f"Section {section_name}: expected {config.expected_records} records, got {len(records)}")
        if config.expected_questions is not None:
            if len(config.expected_questions) != len(records):
                raise ValueError(f"Section {section_name}: expected_questions length {len(config.expected_questions)} != records {len(records)}")
            for i, (record, exp) in enumerate(zip(records, config.expected_questions)):
                if len(record.scores) != exp:
                    raise ValueError(f"Section {section_name}: record {i+1} expected {exp} questions, got {len(record.scores)}")

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

                    record.eval_results = {
                        "sum": total,
                        "mean": mean,
                        "normed_sum": normed_sum
                    }
                    
                    # Get the section-relative row and add section_start_row for absolute position
                    section_relative_row, col_idx = record.details.location
                    row_idx = section_relative_row + section_result.details.start_row
                    
                    # Column U (20) for Sum, Column V (21) for Mean
                    self.df.iloc[row_idx, self.RCIX + 0] = total

                    mean = round(mean, 2)
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
        selection_classification : Dict[str, dict[str, Any]] = {}
        for class_id in set(classification):
            marker = class_marker[class_id - 1]  # assuming class_id starts from 1
            records_in_class = [(i, idx, section_result.records[i]) for i, (c, idx) in enumerate(zip(classification, question_id)) if c == class_id]

            means = [r.eval_results['mean'] for i, idx, r in records_in_class if r.eval_results and 'mean' in r.eval_results]
            std_of_means = np.std(means) if means else 0
            max_of_means = np.max(means) if means else 0

            high_grade_limit = (1.0 - 0.001) * (max_of_means - std_of_means) # 0.001 error margin

            selected_records = []

            for i, idx, record in records_in_class:

                section_relative_row, col_idx = record.details.location
                row_idx = section_relative_row + section_result.details.start_row

                if record.eval_results and 'mean' in record.eval_results:
                    is_high = record.eval_results['mean'] >= high_grade_limit
                    if is_high:
                        selected_records.append(idx)

                    # add to excel
                    self.df.iloc[row_idx, self.RCIX + 5] = "High" if is_high else "Other"

            class_evaluation = f"{marker}: { ', '.join(map(str, selected_records)) }"

            i, idx, first_record = records_in_class[0]
            section_relative_row, col_idx = first_record.details.location
            row_idx = section_relative_row + section_result.details.start_row

            self.df.iloc[row_idx, self.RCIX + 3] = round(high_grade_limit, 2)
            self.df.iloc[row_idx, self.RCIX + 4] = class_evaluation
            selection_classification[marker] = {
                "records": records_in_class,
                "means": means,
                "std_of_means": std_of_means,
                "max_of_means": max_of_means,
                "high_grade_limit": high_grade_limit,
                "class_evaluation": class_evaluation
            }



    def update_excel_with_statistics(self, output_file: str) -> None:
        """
        Calculates sum and mean for all records and updates the Excel file with results in columns S and T.
        Adds column headers at section start rows.
        
        Parameters:
        output_file (str): The path to save the updated Excel file.
        """
        if self.df is None:
            print("No data loaded.")
            return

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

    
