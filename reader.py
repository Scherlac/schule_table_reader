
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re
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
                    if record['name']:
                        records.append(record)
                elif mode == 'multi':
                    if record['name'] and not record['items'] and not record['scores']:
                        if pending:
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
        Parses a string to extract name, item numbers, and their modifiers using a single regex pattern.

        This method uses a single regex with capture groups to find all item numbers and modifiers,
        then extracts the name by removing the matched patterns. It handles all formats uniformly.

        Regex pattern (multiline with named groups for maintainability):
        (?P<num>\d+)     # Capture group for item number (one or more digits)
        (?P<mod>[a-zA-Z]*) # Capture group for optional modifier (zero or more letters)

        Parameters:
        s (str): The string containing the name and item numbers.

        Returns:
        tuple: (name str, list of ints for items, list of strs for modifiers)
        """
        pattern = r'(?P<num>\d+)(?P<mod>[a-zA-Z]*)'
        matches = re.findall(pattern, s)
        name = re.sub(pattern, '', s).strip()
        items = [int(match[0]) for match in matches]
        modifiers = [match[1] for match in matches]
        return name, items, modifiers


class SectionParser:
    """
    Parser class for extracting structured data from Excel sections.
    """

    def __init__(self) -> None:
        self.record_parser = RecordParser()

    def parse_section(self, df: pd.DataFrame, section_name: str, approx_start: int, all_sections: List[str], multi_row: bool = False) -> List[Dict[str, Any]]:
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
        return self.record_parser.parse_records(section_df, mode)


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

        # Parse sections by providing whole data and approximate starts
        all_sections = ['TANSTÍLUS', 'MOTIVÁCIÓ', 'KATT']
        self.parsed_sections['TANSTÍLUS'] = self.parser.parse_section(self.df, 'TANSTÍLUS', 0, all_sections, multi_row=False)
        self.parsed_sections['MOTIVÁCIÓ'] = self.parser.parse_section(self.df, 'MOTIVÁCIÓ', 15, all_sections, multi_row=False)
        self.parsed_sections['KATT'] = self.parser.parse_section(self.df, 'KATT', 25, all_sections, multi_row=True)

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

    def get_full_data(self) -> pd.DataFrame:
        """
        Returns the full DataFrame.

        Returns:
        pd.DataFrame: The full data from the Excel file.
        """
        return self.df

    
