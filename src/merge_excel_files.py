#!/usr/bin/env python3
"""
File Merger - Main Script

This script merges multiple Excel and CSV files from subdirectories into consolidated files.
Usage: python merge_excel_files.py <directory_path>
"""

import sys
import argparse
from pathlib import Path

# Add the src directory to Python path for imports
sys.path.insert(0, str(Path(__file__).parent))

from file_merger import FileMerger


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
    
    args = parser.parse_args()
    
    # Convert to Path object and resolve
    directory_path = Path(args.directory).resolve()
    
    print("File Merger (Excel & CSV)")
    print("=" * 50)
    print(f"Target directory: {directory_path}")
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
        column_cleaning_options=column_cleaning_options
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
