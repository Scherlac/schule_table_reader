# Excel Assessment Processing Tool

A comprehensive tool for processing educational assessment Excel files, generating statistics, and creating evaluation summaries for psychological and educational data.

## Features

- **Batch Processing**: Process multiple Excel files with a single command
- **Statistics Generation**: Automatic calculation of means, sums, and other statistics
- **Evaluation System**: Grade limit analysis and classification
- **File Preservation**: Original files are never modified; processed versions are saved separately
- **Flexible Validation**: Handles non-conforming files with informative warnings
- **Summary Reports**: Combined evaluation data across multiple files
- **CLI Interface**: User-friendly command-line interface with progress tracking

## Installation

### Prerequisites

- Python 3.8 or higher
- Required packages (automatically installed via requirements.txt)

### Setup

1. Clone or download the project
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Command Line Interface

The tool provides a Click-based CLI with two main commands:

#### Process Multiple Files

Process all Excel files in a directory and generate a summary:

```bash
python cli.py process [OPTIONS]
```

**Options:**
- `-i, --input-dir PATH`: Input directory containing Excel files (default: current directory)
- `-o, --output-dir PATH`: Output directory for processed files (default: 'result')
- `-p, --pattern TEXT`: File pattern to match (default: '*.xlsx')
- `-v, --verbose`: Enable verbose output with detailed processing information

**Examples:**
```bash
# Process all Excel files in current directory
python cli.py process

# Process files with custom output directory and verbose output
python cli.py process -o results -v

# Process specific file types from a custom directory
python cli.py process -i data -o output -p "*.xlsx" -v
```

#### Process Single File

Process a single Excel file:

```bash
python cli.py process-single [OPTIONS] FILENAME
```

**Arguments:**
- `FILENAME`: Path to the Excel file to process

**Options:**
- `-o, --output-dir PATH`: Output directory for processed file (default: 'result')
- `-v, --verbose`: Enable verbose output

**Examples:**
```bash
# Process a single file
python cli.py process-single myfile.xlsx

# Process with custom output directory
python cli.py process-single myfile.xlsx -o processed_files -v
```

#### General Options

```bash
# Show help
python cli.py --help

# Show version
python cli.py --version
```

### Output Structure

The tool creates the following outputs:

```
output_directory/
├── filename_processed.xlsx    # Original file with statistics added
├── summary.xlsx              # Combined evaluation data (batch processing only)
└── ...                       # Additional processed files
```

## File Format

### Expected Excel Structure

The tool expects Excel files with the following sections:

1. **TANSTÍLUS** (Learning Styles): 9 records with cognitive/social/logic classifications
2. **RAVEN** (IQ Test): 1 record with IQ score
3. **MOTIVÁCIÓ** (Motivation): 10 records with various motivation dimensions
4. **KATT** (Learning Assessment): 6 records with learning-related metrics

### Flexible Processing

The tool is designed to handle non-conforming files:
- **Missing sections**: Warns and skips evaluation for missing sections
- **Incomplete data**: Processes available data with warnings
- **Extra records**: Accommodates variations in record counts
- **Partial scores**: Uses available scores even if some questions are missing

## Configuration

### Section Configuration

Section definitions and validation rules are stored in `sections_config.json`:

```json
{
  "SECTION_NAME": {
    "approx_start": 0,
    "multi_row": false,
    "expected_records": 9,
    "expected_questions": [4, 4, 8, 9, 6, 6, 6, 10, 4],
    "classification": [1, 1, 1, 1, 1, 2, 2, 3, 3],
    "class_marker": ["cognitive", "social", "logic"],
    "question_id": [11, 12, 21, 22, 31, 41, 42, 51, 52],
    "class_eval": "grade_limit,> grade_limit"
  }
}
```

## Examples

### Basic Usage

```bash
# Process all Excel files in current directory
python cli.py process

# Process with verbose output
python cli.py process -v

# Process files from specific directory
python cli.py process -i ./data -o ./results -v
```

### Single File Processing

```bash
# Process individual file
python cli.py process-single assessment.xlsx -o ./output -v
```

### Batch Processing Output

When processing multiple files, you'll see:
- Progress bar showing processing status
- Warnings for any data inconsistencies
- Final summary with file counts and locations
- Combined evaluation data in `summary.xlsx`

## Troubleshooting

### Common Issues

1. **"No Excel files found"**
   - Check file extensions (must be .xlsx)
   - Verify file pattern matches your files
   - Ensure files are not open in Excel (may create temp files)

2. **"Section X has no records"**
   - This is a warning, not an error
   - The file may be missing that section
   - Processing continues with available data

3. **"Expected X questions, got Y"**
   - File has incomplete data for some records
   - Processing uses available scores
   - Check source data for completeness

4. **Import errors**
   - Ensure all dependencies are installed: `pip install -r requirements.txt`
   - Check Python version (3.8+ required)

### Verbose Mode

Use `-v` flag for detailed processing information:
- Record-by-record processing details
- Score validation information
- Section parsing progress
- Warning details for data issues

## Technical Details

### Dependencies

- `pandas`: Data manipulation and Excel processing
- `openpyxl`: Excel file reading/writing
- `click`: Command-line interface framework
- `pydantic`: Data validation and models
- `numpy`: Statistical calculations

### Architecture

- **CLI Layer** (`cli.py`): Command-line interface and argument parsing
- **Processing Layer** (`reader.py`): Core Excel processing and evaluation logic
- **Models Layer** (`models.py`): Data structures and validation
- **Configuration** (`sections_config.json`): Section definitions and rules

## Sample Data

The repository includes sample Excel files demonstrating the expected format:
- `példatáblázat.xlsx`: Complete assessment file
- `példagyerek2.xlsx`, `példagyerek3.xlsx`: Files with some missing sections

## Contributing

When modifying the code:
1. Update tests in `test_reader.py`
2. Update configuration in `sections_config.json` if adding new sections
3. Update this README if adding new features
4. Test with various file formats and edge cases

## License

MIT License

Copyright (c) 2026 Excel Assessment Processing Tool

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


