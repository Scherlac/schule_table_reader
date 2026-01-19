#!/usr/bin/env python3
"""
Command-line interface for Excel processing and evaluation.
"""

import os
import glob
import click
import pandas as pd
from reader import ExcelImporter


@click.group()
@click.version_option(version="1.0.0")
def cli():
    """Excel Processing and Evaluation CLI Tool.

    Process Excel files containing educational assessment data,
    generate statistics, and create evaluation summaries.
    """
    pass


@cli.command()
@click.option(
    '--input-dir',
    '-i',
    default='.',
    help='Directory containing Excel files to process (default: current directory)'
)
@click.option(
    '--output-dir',
    '-o',
    default='result',
    help='Directory to save processed files and summary (default: result)'
)
@click.option(
    '--pattern',
    '-p',
    default='*.xlsx',
    help='File pattern to match Excel files (default: *.xlsx)'
)
@click.option(
    '--skip-summary',
    is_flag=True,
    help='Skip creating the summary.xlsx file'
)
@click.option(
    '--verbose',
    '-v',
    is_flag=True,
    help='Enable verbose output'
)
def process(input_dir, output_dir, pattern, skip_summary, verbose):
    """Process Excel files and generate evaluation summaries.

    This command processes all Excel files in the input directory,
    generates statistics for each file, and creates a summary
    with concatenated evaluation data.
    """
    # Convert to absolute paths
    input_dir = os.path.abspath(input_dir)
    output_dir = os.path.abspath(output_dir)

    if verbose:
        click.echo(f"Input directory: {input_dir}")
        click.echo(f"Output directory: {output_dir}")
        click.echo(f"File pattern: {pattern}")

    # Change to input directory to find files
    original_cwd = os.getcwd()
    try:
        os.chdir(input_dir)

        # Find all Excel files matching the pattern
        excel_files = glob.glob(pattern)
        excel_files = [f for f in excel_files if not f.startswith('~$')]  # Exclude temp files

        if not excel_files:
            click.echo(f"No Excel files found matching pattern '{pattern}' in {input_dir}")
            return

        click.echo(f"Found {len(excel_files)} Excel files: {', '.join(excel_files)}")

        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            click.echo(f"Created output directory: {output_dir}")

        # List to collect evaluation DataFrames
        evaluation_dfs = []

        # Process each Excel file
        with click.progressbar(excel_files, label='Processing files') as files:
            for file_path in files:
                try:
                    if verbose:
                        click.echo(f"\nProcessing: {file_path}")

                    # Create importer instance
                    importer = ExcelImporter(file_path)

                    if importer.df is None:
                        click.echo(f"Warning: Failed to load Excel file: {file_path}", err=True)
                        continue

                    if verbose:
                        # Dump the report
                        importer.dump_report()

                    # Update Excel with statistics (output to result folder)
                    base_name = os.path.splitext(file_path)[0]
                    output_file = os.path.join(output_dir, f'{base_name}_processed.xlsx')
                    importer.update_excel_with_statistics(output_file)

                    # Get evaluation data
                    df_eval = importer.evaluate()
                    if df_eval is not None and not df_eval.empty:
                        # Add source file column
                        df_eval['source_file'] = file_path
                        evaluation_dfs.append(df_eval)
                        if verbose:
                            click.echo(f"Collected evaluation data with {len(df_eval.columns)} columns")
                    else:
                        if verbose:
                            click.echo("No evaluation data collected")

                except Exception as e:
                    click.echo(f"Error processing {file_path}: {e}", err=True)
                    continue

        # Concatenate all evaluation DataFrames
        if evaluation_dfs and not skip_summary:
            if verbose:
                click.echo(f"\nConcatenating {len(evaluation_dfs)} evaluation DataFrames...")

            summary_df = pd.concat(evaluation_dfs, ignore_index=False)

            # Save summary to Excel
            summary_path = os.path.join(output_dir, 'summary.xlsx')
            summary_df.to_excel(summary_path)

            click.echo(f"Summary saved to: {summary_path}")
            click.echo(f"Summary shape: {summary_df.shape} (rows x columns)")

            if verbose:
                click.echo(f"Summary columns: {', '.join(summary_df.columns)}")
                click.echo("\nFirst few rows of summary:")
                click.echo(summary_df.head().to_string())
        elif skip_summary:
            click.echo("Skipped summary creation as requested.")
        else:
            click.echo("No evaluation data collected from any files.")

    finally:
        # Restore original working directory
        os.chdir(original_cwd)


@cli.command()
@click.argument('file_path')
@click.option(
    '--output-dir',
    '-o',
    default='result',
    help='Directory to save processed file (default: result)'
)
@click.option(
    '--verbose',
    '-v',
    is_flag=True,
    help='Enable verbose output'
)
def process_single(file_path, output_dir, verbose):
    """Process a single Excel file.

    FILE_PATH: Path to the Excel file to process
    """
    if not os.path.exists(file_path):
        click.echo(f"Error: File '{file_path}' does not exist", err=True)
        return

    output_dir = os.path.abspath(output_dir)

    if verbose:
        click.echo(f"Processing file: {file_path}")
        click.echo(f"Output directory: {output_dir}")

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        click.echo(f"Created output directory: {output_dir}")

    try:
        # Create importer instance
        importer = ExcelImporter(file_path)

        if importer.df is None:
            click.echo(f"Error: Failed to load Excel file: {file_path}", err=True)
            return

        if verbose:
            # Dump the report
            importer.dump_report()

        # Update Excel with statistics
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_file = os.path.join(output_dir, f'{base_name}_processed.xlsx')
        importer.update_excel_with_statistics(output_file)

        # Get and display evaluation data
        df_eval = importer.evaluate()
        if df_eval is not None and not df_eval.empty:
            click.echo(f"Evaluation data shape: {df_eval.shape}")
            if verbose:
                click.echo("Evaluation data:")
                click.echo(df_eval.to_string())
        else:
            click.echo("No evaluation data generated.")

        click.echo(f"Processed file saved to: {output_file}")

    except Exception as e:
        click.echo(f"Error processing file: {e}", err=True)


if __name__ == '__main__':
    cli()