#!/usr/bin/env python3
"""
File Merger - Main Script

This script merges multiple Excel and CSV files from subdirectories into consolidated files.
Usage: python merge_excel_files.py <directory_path>
"""

import sys
import argparse
from pathlib import Path
from typing import List

# Add the src directory to Python path for imports
sys.path.insert(0, str(Path(__file__).parent))

from file_merger import FileMerger


def load_columns_from_yaml(path: Path) -> List[str]:
    """
    Load column names from a YAML file with a top-level 'columns' list.
    Raises ValueError if the file is missing or invalid.
    """
    import yaml
    if not path.exists():
        raise ValueError(f"Columns file not found: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    if not data or 'columns' not in data:
        raise ValueError(f"YAML file {path} must contain a 'columns' key with a list")
    cols = data['columns']
    if not isinstance(cols, list):
        raise ValueError(f"'columns' in {path} must be a list")
    return [str(c).strip() for c in cols if c]


def main():
    """Main function to handle command-line interface."""
    parser = argparse.ArgumentParser(
        description="Merge Excel and CSV files from subdirectories into consolidated files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python merge_excel_files.py /path/to/data
  python merge_excel_files.py ./Data
  python merge_excel_files.py "C:\\Users\\Data"
  python merge_excel_files.py ./Data --lowercase-columns
  python merge_excel_files.py ./Data --no-clean-columns
  python merge_excel_files.py ./Data --remove-special-chars --verbose
        """
    )
    
    parser.add_argument(
        'directory',
        help='Path to the root directory containing subdirectories with Excel and CSV files'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose output'
    )
    
    parser.add_argument(
        '--no-clean-columns',
        action='store_true',
        help='Disable column name cleaning'
    )
    
    parser.add_argument(
        '--lowercase-columns',
        action='store_true',
        help='Convert column names to lowercase'
    )
    
    parser.add_argument(
        '--remove-special-chars',
        action='store_true',
        help='Remove special characters from column names'
    )
    
    parser.add_argument(
        '--no-validate-structure',
        action='store_true',
        help='Skip column structure validation (merge even if column sets differ)'
    )
    
    parser.add_argument(
        '--use-columns',
        action='store_true',
        help='Use column list from a YAML file; only these columns are read and exported'
    )
    
    parser.add_argument(
        '--columns-file',
        type=str,
        default=None,
        metavar='PATH',
        help='Path to YAML file with "columns:" list (default: columns.yaml in current directory)'
    )
    
    args = parser.parse_args()
    
    # Convert to Path object and resolve
    directory_path = Path(args.directory).resolve()
    
    # Load columns from YAML when --use-columns is passed
    columns_to_export = None
    if args.use_columns:
        columns_path = Path(args.columns_file).resolve() if args.columns_file else Path.cwd() / 'columns.yaml'
        try:
            columns_to_export = load_columns_from_yaml(columns_path)
            if not columns_to_export:
                print(f"Error: columns file {columns_path} has an empty 'columns' list.")
                sys.exit(2)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(2)
    
    print("File Merger (Excel & CSV)")
    print("=" * 50)
    print(f"Target directory: {directory_path}")
    if columns_to_export:
        print(f"Columns filter: {len(columns_to_export)} columns from YAML")
    print()
    
    # Configure column cleaning options
    column_cleaning_options = {
        'lowercase': args.lowercase_columns,
        'remove_special_chars': args.remove_special_chars
    }
    
    # Create file merger instance
    merger = FileMerger(
        str(directory_path),
        clean_columns=not args.no_clean_columns,
        column_cleaning_options=column_cleaning_options,
        validate_structure=not args.no_validate_structure,
        columns_to_export=columns_to_export
    )
    
    # Process all subdirectories
    results = merger.process_all_subdirectories()
    
    # Print summary
    merger.print_summary(results)
    
    # Exit with appropriate code
    if results['success'] and not results['errors']:
        sys.exit(0)
    elif results['success'] and results['errors']:
        sys.exit(1)  # Partial success
    else:
        sys.exit(2)  # Complete failure


if __name__ == "__main__":
    main()
