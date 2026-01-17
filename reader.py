
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re
import json
from typing import List, Dict, Tuple, Optional, Any


def read_excel_file(file_path: str) -> pd.DataFrame:
    """
    Reads an Excel file and returns its contents as a pandas DataFrame.

    Parameters:
    file_path (str): The path to the Excel file.

    Returns:
    pd.DataFrame: DataFrame containing the data from the Excel file.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None


class RecordParser:
    """
    Parser class for extracting structured data from individual Excel rows and sections.
    """

    def __init__(self) -> None:
        pass

    def parse_records(self, df: pd.DataFrame, mode: str) -> List[Dict[str, Any]]:
        """
        Parses the DataFrame into a list of structured records, handling single or multi-row records.

        Parameters:
        df (pd.DataFrame): The section DataFrame.
        mode (str): 'single' for single-row records, 'multi' for multi-row records.

        Returns:
        list: List of dicts, each with 'name', 'items', 'modifiers', 'scores'.
        """
        records = []
        pending = None
        for idx, row in df.iterrows():
            record = self.parse_record(row)
            if record:
                if mode == 'single':
                    if record['name'] and record['scores']:
                        records.append(record)
                elif mode == 'multi':
                    if record['name'] and not record['items'] and not record['scores']:
                        if pending and pending.get('scores'):
                            records.append(pending)
                        pending = record
                    elif record['items'] and not record['name'] and not record['scores']:
                        if pending:
                            pending['items'] = record['items']
                            pending['modifiers'] = record['modifiers']
                    elif record['scores']:
                        if pending:
                            pending['scores'] = record['scores']
                            records.append(pending)
                            pending = None
        if mode == 'multi' and pending:
            records.append(pending)
        return records

    def parse_record(self, row) -> Optional[Dict[str, Any]]:
        """
        Parses a single row into a structured record.

        Parameters:
        row: The DataFrame row.

        Returns:
        dict: Dictionary with 'label', 'name', 'items', 'modifiers', 'scores' or None if no label.
        """
        label = row[3] if len(row) > 3 and pd.notna(row[3]) else None
        if label:
            name, items, modifiers = self._parse_item_string(str(label))
            scores = [row[i] for i in range(4, len(row)) if pd.notna(row[i])]
            return {
                'label': str(label),
                'name': name,
                'items': items,
                'modifiers': modifiers,
                'scores': scores
            }
        return None

    def _parse_item_string(self, s: str) -> Tuple[str, List[int], List[str]]:
        """
        Parses a string to extract name, item numbers, and their modifiers.

        Handles different formats: space-separated, comma-separated, semicolon-separated.

        Parameters:
        s (str): The string containing the name and item numbers.

        Returns:
        tuple: (name str, list of ints for items, list of strs for modifiers)
        """
        if ';' in s:
            name = re.sub(r'\d+[a-zA-Z]*', '', s).strip()
            parts = []
            for x in s.split(';'):
                x = x.strip()
                if x:
                    if ',' in x:
                        parts.extend([p.strip() for p in x.split(',') if p.strip()])
                    else:
                        parts.append(x)
            items = []
            modifiers = []
            for p in parts:
                match = re.match(r'(\d+)([a-zA-Z]*)', p)
                if match:
                    items.append(int(match.group(1)))
                    modifiers.append(match.group(2))
            return name, items, modifiers
        else:
            parts = s.split()
            item_parts = []
            name_parts = []
            for p in parts:
                if re.match(r'\d+[a-zA-Z]*', p):
                    item_parts.append(p)
                else:
                    name_parts.append(p)
            name = ' '.join(name_parts)
            items = []
            modifiers = []
            for p in item_parts:
                match = re.match(r'(\d+)([a-zA-Z]*)', p)
                if match:
                    items.append(int(match.group(1)))
                    modifiers.append(match.group(2))
            return name, items, modifiers


class SectionParser:
    """
    Parser class for extracting structured data from Excel sections.
    """

    def __init__(self) -> None:
        self.record_parser = RecordParser()

    def parse_section(self, df: pd.DataFrame, section_name: str, approx_start: int, all_sections: List[str], multi_row: bool = False, expected_records: Optional[int] = None, expected_questions: Optional[List[int]] = None) -> List[Dict[str, Any]]:
        """
        Parses the section DataFrame into a list of structured records.

        Parameters:
        df (pd.DataFrame): The full DataFrame.
        section_name (str): The name of the section to parse.
        approx_start (int): Approximate row to start searching for the section.
        all_sections (List[str]): List of all possible section names.
        multi_row (bool): Whether records span multiple rows.

        Returns:
        list: List of dicts, each with 'name', 'items', 'modifiers', 'scores'.
        """
        # Find the exact start and end rows
        start_row = None
        end_row = len(df)
        markers = [s for s in all_sections if s != section_name]
        for idx in range(approx_start, len(df)):
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
            return []

        section_df = df.iloc[start_row:end_row]

        mode = 'multi' if multi_row else 'single'
        records = self.record_parser.parse_records(section_df, mode)
        if expected_records is not None and len(records) != expected_records:
            raise ValueError(f"Section {section_name}: expected {expected_records} records, got {len(records)}")
        if expected_questions is not None:
            if len(expected_questions) != len(records):
                raise ValueError(f"Section {section_name}: expected_questions length {len(expected_questions)} != records {len(records)}")
            for i, (record, exp) in enumerate(zip(records, expected_questions)):
                if len(record['scores']) != exp:
                    raise ValueError(f"Section {section_name}: record {i+1} expected {exp} questions, got {len(record['scores'])}")
        return records


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
        try:
            self.df = pd.read_excel(file_path, header=None)
            self.sections = {}
            self.parsed_sections = {}
            self.child_name = None
            self.parser = SectionParser()
            self._extract_data()
        except Exception as e:
            print(f"An error occurred while importing the Excel file: {e}")
            self.df = None

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
            sections_config = json.load(f)

        all_sections = list(sections_config.keys())

        # Parse sections using configuration
        for section_name, config in sections_config.items():
            self.parsed_sections[section_name] = self.parser.parse_section(
                self.df,
                section_name,
                config['approx_start'],
                all_sections,
                multi_row=config['multi_row'],
                expected_records=config.get('expected_records'),
                expected_questions=config.get('expected_questions')
            )

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

    def get_parsed_sections(self) -> Dict[str, List[Dict[str, Any]]]:
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

        # Full data summary
        full_data = self.get_full_data()
        print(f"\nFull Data Shape: {full_data.shape}")
        print("First 5 rows of full data:")
        print(full_data.head().to_string(index=False))

    def get_full_data(self) -> pd.DataFrame:
        """
        Returns the full DataFrame.

        Returns:
        pd.DataFrame: The full data from the Excel file.
        """
        return self.df

    
