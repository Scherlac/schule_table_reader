
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re
import json
import enum
from typing import List, Dict, Tuple, Optional, Any
from pydantic import BaseModel, Field, ConfigDict


class SectionDetails(BaseModel):
    selection_df: Any = Field(description="The DataFrame containing the selected section data")
    start_row: int = Field(description="The starting row index of the section")
    end_row: int = Field(description="The ending row index of the section")


class RecordDetails(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)
    
    source: pd.DataFrame = Field(description="The DataFrame containing the source data")
    location: Tuple[int, int] = Field(description="The (row, column) coordinates in the table")


class Record(BaseModel):
    label: str = Field(description="The original label string from the Excel row")
    name: str = Field(description="The parsed name of the record")
    items: List[int] = Field(description="List of item numbers associated with the record")
    modifiers: List[str] = Field(description="List of modifiers for each item")
    scores: List[Any] = Field(description="List of score values for the record")
    subsection: Optional[str] = Field(default=None, description="Optional subsection name if the record belongs to a subsection")
    details: Optional[RecordDetails] = Field(default=None, description="Details about the source and location of the data")


class REPORT_EVAL_ENUM(enum.Flag):
    HIGH_GRADES = enum.auto()
    NORMED_SUM = enum.auto()
    MEAN = enum.auto()
    COPY = enum.auto()

report_evlal_dict = {
    "> max - dev": REPORT_EVAL_ENUM.HIGH_GRADES,
    "normed sum": REPORT_EVAL_ENUM.NORMED_SUM,
    "mean": REPORT_EVAL_ENUM.MEAN,
    "copy": REPORT_EVAL_ENUM.COPY
}


class SectionConfig(BaseModel):
    approx_start: int = Field(description="Approximate starting row for searching the section")
    multi_row: bool = Field(description="Whether records in this section span multiple rows")
    expected_records: Optional[int] = Field(default=None, description="Expected number of records in the section")
    expected_questions: Optional[List[int]] = Field(default=None, description="Expected number of questions per record")
    # "classification": [1, 1, 1, 1, 1, 2, 2, 3, 3],
    classification: Optional[List[int]] = Field(default=None, description="Class identifiers for each record"),
    # "class_marker": ["cognitive", "social", "logic"],
    class_marker: Optional[List[str]] = Field(default=None, description="Class markers for grouping records"),
    # "question_id": [11, 12, 21, 22, 31, 41, 42, 51, 52],
    question_id: Optional[List[int]] = Field(default=None, description="Question IDs for each record"),
    # "report_eval": "> max - dev"
    report_eval: Optional[str] = Field(default=None, description="Evaluation criteria for reporting scores")

    @property
    def report_eval_flags(self) -> REPORT_EVAL_ENUM:
        """
        Parses the report_eval string into a combined REPORT_EVAL_ENUM flag.

        Returns:
        REPORT_EVAL_ENUM: Combined flags based on the report_eval configuration.
        """
        if not self.report_eval:
            return REPORT_EVAL_ENUM(0)
        flags = REPORT_EVAL_ENUM(0)
        parts = [part.strip() for part in self.report_eval.split(',')]
        for part in parts:
            if part in report_evlal_dict:
                flags |= report_evlal_dict[part]
        return flags
    
    @report_eval_flags.setter
    def report_eval_flags(self, flags: REPORT_EVAL_ENUM) -> None:
        """
        Sets the report_eval string based on the provided REPORT_EVAL_ENUM flags.

        Parameters:
        flags (REPORT_EVAL_ENUM): Combined flags to set the report_eval configuration.
        """
        parts = []
        for key, value in report_evlal_dict.items():
            if flags & value:
                parts.append(key)
        self.report_eval = ', '.join(parts)



class SectionResult(BaseModel):
    details: SectionDetails = Field(description="Details about the section extraction")
    records: List[Record] = Field(description="List of parsed records from the section")
    # subsections: Optional[Dict[str, List[Record]]] = Field(default=None, description="Optional mapping of subsection names to their records")


class RecordParser:
    """
    Parser class for extracting structured data from individual Excel rows and sections.
    """

    def __init__(self) -> None:
        pass

    def parse_records(self, df: pd.DataFrame, section_name: str, mode: str) -> List[Record]:
        """
        Parses the DataFrame into a list of structured records, handling single or multi-row records.

        Parameters:
        df (pd.DataFrame): The section DataFrame.
        mode (str): 'single' for single-row records, 'multi' for multi-row records.
        section_name (str): the name of the section

        Returns:
        list: List of Record objects, each with 'name', 'items', 'modifiers', 'scores'.
        """
        records = []
        pending = None
        subsections = None
        for idx, row in df.iterrows():
            record = self.parse_record(row)
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

    def parse_record(self, row) -> Optional[Record]:
        """
        Parses a single row into a structured record.

        Parameters:
        row: The DataFrame row.

        Returns:
        Record: Record object with 'label', 'name', 'items', 'modifiers', 'scores' or None if no label.
        """
        if len(row) > 3 and pd.notna(row[3]):
            label = row[3]
            name, items, modifiers = self._parse_item_string(str(label))
            scores = [row[i] for i in range(4, len(row)) if pd.notna(row[i])]
            return Record(
                label=str(label),
                name=name,
                items=items,
                modifiers=modifiers,
                scores=scores
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
        column_index = -1  # Search all columns
        markers = [s for s in self.all_sections if s != section_name]
        for idx in range(self.config.approx_start, len(df)):
            row = df.iloc[idx]
            for cell in row:
                if pd.notna(cell):
                    cell_str = str(cell)
                    if start_row is None and section_name in cell_str:
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
        records = self.record_parser.parse_records(section_df, section_name, mode)
        self._validate_section_records(section_name, records, self.config)
        return SectionResult(
            details=SectionDetails(selection_df=section_df, start_row=start_row, end_row=end_row),
            records=records
        )


class ExcelImporter:
    """
    Importer class to capture and structure the content of the Excel file.
    """

    def __init__(self, file_path: str) -> None:
        """
        Initializes the importer by reading the Excel file.

        Parameters:
        file_path (str): The path to the Excel file.
        """
        # try:
        if True:
            self.df = pd.read_excel(file_path, header=None)
            self.sections = {}
            self.parsed_sections = {}
            self.child_name = None
            self._extract_data()
        # except Exception as e:
        #     print(f"An error occurred while importing the Excel file: {e}")
        #     self.df = None

    def _extract_data(self) -> None:
        """
        Extracts the child name and sections from the DataFrame.
        """
        if self.df is None:
            return

        # Extract child name (assuming in row 1, column 3 - 0-indexed row 1, col 3)
        if len(self.df) > 1 and len(self.df.columns) > 3:
            self.child_name = self.df.iloc[1, 3] if pd.notna(self.df.iloc[1, 3]) else None

        # Load section configuration from JSON
        with open('sections_config.json', 'r', encoding='utf-8') as f:
            raw_config = json.load(f)
            sections_config = {k: SectionConfig.model_validate(v) for k, v in raw_config.items()}

        all_sections = list(sections_config.keys())

        # Parse sections using configuration
        for section_name, config in sections_config.items():
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
                    total = sum(scores)
                    mean = total / count if count > 0 else 0
                    print(f"  Record {i+1} ({display_name}): Count={count}, Sum={total}, Mean={mean:.2f}")
                else:
                    print(f"  Record {i+1} ({display_name}): No scores")

    #     # Full data summary
    #     full_data = self.get_full_data()
    #     print(f"\nFull Data Shape: {full_data.shape}")
    #     print("First 5 rows of full data:")
    #     print(full_data.head().to_string(index=False))

    # def get_full_data(self) -> pd.DataFrame:
    #     """
    #     Returns the full DataFrame.

    #     Returns:
    #     pd.DataFrame: The full data from the Excel file.
    #     """
    #     return self.df

    
