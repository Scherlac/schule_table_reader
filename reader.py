
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
    Parser class for extracting structured data from individual Excel rows.
    """

    def __init__(self) -> None:
        pass

    def parse_record(self, row) -> Optional[Dict[str, Any]]:
        """
        Parses a single row into a structured record.

        Parameters:
        row: The DataFrame row.

        Returns:
        dict: Dictionary with 'name', 'items', 'modifiers', 'scores' or None if no label.
        """
        label = row[3] if len(row) > 3 and pd.notna(row[3]) else None
        if label:
            name, items, modifiers = self._parse_item_string(str(label))
            scores = [row[i] for i in range(4, len(row)) if pd.notna(row[i])]
            return {
                'name': name,
                'items': items,
                'modifiers': modifiers,
                'scores': scores
            }
        return None

    def _parse_item_string(self, s: str) -> Tuple[str, List[int], List[str]]:
        """
        Parses a string to extract name, item numbers, and their modifiers.

        Parameters:
        s (str): The string containing the name and item numbers.

        Returns:
        tuple: (name str, list of ints for items, list of strs for modifiers)
        """
        if ';' in s:
            name = ""
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
            item_parts = [p for p in parts if re.match(r'\d', p)]
            items = []
            modifiers = []
            for p in item_parts:
                match = re.match(r'(\d+)([a-zA-Z]*)', p)
                if match:
                    items.append(int(match.group(1)))
                    modifiers.append(match.group(2))
            name = s
            for p in item_parts:
                name = name.replace(p, '')
            name = ' '.join(name.split())
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
        # Find the exact start row
        start_row = None
        for idx in range(approx_start, len(df)):
            row = df.iloc[idx]
            for cell in row:
                if pd.notna(cell) and section_name in str(cell):
                    start_row = idx
                    break
            if start_row is not None:
                break

        if start_row is None:
            return []

        # Find the end row (next section or end)
        end_row = len(df)
        markers = [s for s in all_sections if s != section_name]
        for idx in range(start_row + 1, len(df)):
            row = df.iloc[idx]
            for cell in row:
                if pd.notna(cell):
                    for marker in markers:
                        if marker in str(cell):
                            end_row = idx
                            break
            if end_row != len(df):
                break

        section_df = df.iloc[start_row:end_row]

        if multi_row:
            return self._parse_multi_row_section(section_df)
        else:
            return self._parse_single_row_section(section_df)

    def _parse_single_row_section(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """
        Parses sections where each record is in a single row.
        """
        records = []
        for idx, row in df.iterrows():
            record = self.record_parser.parse_record(row)
            if record and record['name']:
                records.append(record)
        return records

    def _parse_multi_row_section(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """
        Parses sections where records may span multiple rows.
        """
        records = []
        pending_name = None
        for idx, row in df.iterrows():
            record = self.record_parser.parse_record(row)
            if record:
                if record['items']:
                    if pending_name and not record['name']:
                        record['name'] = pending_name
                        pending_name = None
                    if record['name']:
                        records.append(record)
                else:
                    if record['name']:
                        if pending_name:
                            records.append({
                                'name': pending_name,
                                'items': [],
                                'modifiers': [],
                                'scores': []
                            })
                        pending_name = record['name']
        if pending_name:
            records.append({
                'name': pending_name,
                'items': [],
                'modifiers': [],
                'scores': []
            })
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

        # Parse sections by providing whole data and approximate starts
        all_sections = ['TANSTÍLUS', 'MOTIVÁCIÓ', 'KATT']
        self.parsed_sections['TANSTÍLUS'] = self.parser.parse_section(self.df, 'TANSTÍLUS', 3, all_sections, multi_row=False)
        self.parsed_sections['MOTIVÁCIÓ'] = self.parser.parse_section(self.df, 'MOTIVÁCIÓ', 22, all_sections, multi_row=False)
        self.parsed_sections['KATT'] = self.parser.parse_section(self.df, 'KATT', 39, all_sections, multi_row=True)

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

    
