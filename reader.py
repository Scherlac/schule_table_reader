
# preprocessing excel file to read relevant data the a structured format
import pandas as pd
import re


def read_excel_file(file_path):
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


class ExcelImporter:
    """
    Importer class to capture and structure the content of the Excel file.
    """

    def __init__(self, file_path):
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
            self._extract_data()
        except Exception as e:
            print(f"An error occurred while importing the Excel file: {e}")
            self.df = None

    def _extract_data(self):
        """
        Extracts the child name and sections from the DataFrame.
        """
        if self.df is None:
            return

        # Extract child name (assuming in row 1, column 3 - 0-indexed row 1, col 3)
        if len(self.df) > 1 and len(self.df.columns) > 3:
            self.child_name = self.df.iloc[1, 3] if pd.notna(self.df.iloc[1, 3]) else None

        # Define section markers and their rows
        markers = ['TANSTÍLUS', 'MOTIVÁCIÓ', 'KATT']
        marker_rows = {}
        for idx, row in self.df.iterrows():
            for cell in row:
                if pd.notna(cell):
                    for marker in markers:
                        if marker in str(cell):
                            marker_rows[marker] = idx

        # Extract sections based on markers
        sorted_markers = sorted(marker_rows.items(), key=lambda x: x[1])
        for i, (marker, start_row) in enumerate(sorted_markers):
            end_row = sorted_markers[i+1][1] if i+1 < len(sorted_markers) else len(self.df)
            section_data = self.df.iloc[start_row:end_row]
            self.sections[marker] = section_data
            self.parsed_sections[marker] = self._parse_section(section_data)

    def _parse_section(self, df):
        """
        Parses the section DataFrame into a list of structured records.

        Parameters:
        df (pd.DataFrame): The section DataFrame.

        Returns:
        list: List of dicts, each with 'name', 'items', 'modifiers', 'scores'.
        """
        records = []
        for idx, row in df.iterrows():
            label = row[3] if len(row) > 3 and pd.notna(row[3]) else None
            if label:
                name, items, modifiers = self._parse_item_string(str(label))
                scores = [row[i] for i in range(4, len(row)) if pd.notna(row[i])]
                records.append({
                    'name': name,
                    'items': items,
                    'modifiers': modifiers,
                    'scores': scores
                })
        return records

    def _parse_item_string(self, s):
        """
        Parses a string to extract name, item numbers, and their modifiers.

        Parameters:
        s (str): The string containing the name and item numbers.

        Returns:
        tuple: (name str, list of ints for items, list of strs for modifiers)
        """
        if ';' in s:
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
            name = s
            for p in parts:
                name = name.replace(p, '')
            name = name.strip()
            if name.endswith(';'):
                name = name[:-1].strip()
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

    def get_child_name(self):
        """
        Returns the child name.

        Returns:
        str: The child name or None if not found.
        """
        return self.child_name

    def get_sections(self):
        """
        Returns the sections data.

        Returns:
        dict: Dictionary with section names as keys and DataFrames as values.
        """
        return self.sections

    def get_parsed_sections(self):
        """
        Returns the parsed sections data.

        Returns:
        dict: Dictionary with section names as keys and parsed data as values.
        """
        return self.parsed_sections

    def get_full_data(self):
        """
        Returns the full DataFrame.

        Returns:
        pd.DataFrame: The full data from the Excel file.
        """
        return self.df

    
