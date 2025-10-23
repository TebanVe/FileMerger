"""
File Merger - Core Module

This module contains the core logic for merging Excel and CSV files from subdirectories.
"""

import os
import glob
import pandas as pd
from pathlib import Path
from typing import List, Dict, Tuple, Optional


class FileMerger:
    """Handles merging of Excel and CSV files within subdirectories."""
    
    def __init__(self, root_directory: str, clean_columns: bool = True, 
                 column_cleaning_options: Dict[str, bool] = None):
        """
        Initialize the Excel merger.
        
        Args:
            root_directory: Path to the root directory containing subdirectories
            clean_columns: Whether to clean column names
            column_cleaning_options: Dictionary of cleaning options
        """
        self.root_directory = Path(root_directory)
        self.processed_subdirs = 0
        self.total_files_merged = 0
        self.skipped_single_file = 0
        self.errors = []
        self.clean_columns = clean_columns
        
        # Default column cleaning options
        self.column_cleaning_options = {
            'strip_whitespace': True,
            'normalize_spaces': True,
            'lowercase': False,
            'remove_special_chars': False,
            'handle_common_variations': True
        }
        
        # Update with user-provided options
        if column_cleaning_options:
            self.column_cleaning_options.update(column_cleaning_options)
        
    def validate_directory(self) -> bool:
        """
        Validate that the root directory exists and contains subdirectories.
        
        Returns:
            bool: True if valid, False otherwise
        """
        if not self.root_directory.exists():
            self.errors.append(f"Directory does not exist: {self.root_directory}")
            return False
            
        if not self.root_directory.is_dir():
            self.errors.append(f"Path is not a directory: {self.root_directory}")
            return False
            
        # Check if there are any subdirectories
        subdirs = [d for d in self.root_directory.iterdir() if d.is_dir()]
        if not subdirs:
            self.errors.append(f"No subdirectories found in: {self.root_directory}")
            return False
            
        return True
    
    def get_supported_files(self, directory: Path) -> List[Path]:
        """
        Get all supported files (Excel and CSV) in a directory.
        
        Args:
            directory: Directory to search for files
            
        Returns:
            List of supported file paths
        """
        supported_extensions = ['*.xlsx', '*.xls', '*.csv']
        supported_files = []
        
        for extension in supported_extensions:
            supported_files.extend(directory.glob(extension))
            
        return supported_files
    
    def _detect_file_format(self, file_path: Path) -> str:
        """
        Detect file format based on extension.
        
        Args:
            file_path: Path to the file
            
        Returns:
            'excel' for .xlsx/.xls files, 'csv' for .csv files
        """
        file_ext = file_path.suffix.lower()
        
        if file_ext in ['.xlsx', '.xls']:
            return 'excel'
        elif file_ext == '.csv':
            return 'csv'
        else:
            return 'unknown'
    
    def clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean column names in a DataFrame.
        
        Args:
            df: DataFrame to clean
            
        Returns:
            DataFrame with cleaned column names
        """
        if not self.clean_columns or df.empty:
            return df
            
        original_columns = df.columns.tolist()
        cleaned_columns = []
        
        for col in original_columns:
            cleaned_col = str(col)
            
            # Strip leading and trailing whitespace
            if self.column_cleaning_options.get('strip_whitespace', True):
                cleaned_col = cleaned_col.strip()
            
            # Normalize multiple spaces to single space
            if self.column_cleaning_options.get('normalize_spaces', True):
                import re
                cleaned_col = re.sub(r'\s+', ' ', cleaned_col)
            
            # Convert to lowercase
            if self.column_cleaning_options.get('lowercase', False):
                cleaned_col = cleaned_col.lower()
            
            # Remove special characters
            if self.column_cleaning_options.get('remove_special_chars', False):
                import re
                cleaned_col = re.sub(r'[^\w\s-]', '', cleaned_col)
            
            # Handle common variations
            if self.column_cleaning_options.get('handle_common_variations', True):
                # Common replacements
                variations = {
                    'id': 'ID',
                    'date': 'Date',
                    'name': 'Name',
                    'email': 'Email',
                    'phone': 'Phone',
                    'address': 'Address',
                    'city': 'City',
                    'state': 'State',
                    'zip': 'ZIP',
                    'zipcode': 'ZIP',
                    'zip_code': 'ZIP'
                }
                
                for old, new in variations.items():
                    if cleaned_col.lower() == old.lower():
                        cleaned_col = new
                        break
            
            cleaned_columns.append(cleaned_col)
        
        # Create new DataFrame with cleaned column names
        df_cleaned = df.copy()
        df_cleaned.columns = cleaned_columns
        
        # Log column changes if any
        changes = []
        for orig, cleaned in zip(original_columns, cleaned_columns):
            if orig != cleaned:
                changes.append(f"'{orig}' → '{cleaned}'")
        
        if changes:
            print(f"    📝 Column names cleaned: {', '.join(changes)}")
        
        return df_cleaned
    
    def read_file(self, file_path: Path) -> Optional[pd.DataFrame]:
        """
        Read a supported file (Excel or CSV) and return a DataFrame.
        
        Args:
            file_path: Path to the file
            
        Returns:
            DataFrame or None if reading fails
        """
        file_format = self._detect_file_format(file_path)
        
        if file_format == 'excel':
            return self.read_excel_file(file_path)
        elif file_format == 'csv':
            return self._read_with_csv(file_path)
        else:
            self.errors.append(f"Unsupported file format: {file_path}")
            return None
    
    def read_excel_file(self, file_path: Path) -> Optional[pd.DataFrame]:
        """
        Read an Excel file and return a DataFrame.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            DataFrame or None if reading fails
        """
        # Try multiple methods in order of preference
        methods = [
            ("pandas with engine detection", self._read_with_pandas),
            ("xlwings", self._read_with_xlwings)
        ]
        
        for method_name, method_func in methods:
            try:
                df = method_func(file_path)
                if df is not None:
                    return df
            except Exception as e:
                # Continue to next method
                continue
        
        # If all methods failed, add error
        self.errors.append(f"Error reading {file_path}: All reading methods failed")
        return None
    
    def read_file_with_method(self, file_path: Path) -> Tuple[Optional[pd.DataFrame], str]:
        """
        Read a supported file and return a DataFrame with the method used.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Tuple of (DataFrame or None, method_name)
        """
        file_format = self._detect_file_format(file_path)
        
        if file_format == 'excel':
            return self.read_excel_file_with_method(file_path)
        elif file_format == 'csv':
            try:
                df = self._read_with_csv(file_path)
                if df is not None:
                    return df, "pandas CSV reader"
                else:
                    return None, "failed"
            except Exception as e:
                self.errors.append(f"Error reading CSV {file_path}: {str(e)}")
                return None, "failed"
        else:
            self.errors.append(f"Unsupported file format: {file_path}")
            return None, "failed"
    
    def read_excel_file_with_method(self, file_path: Path) -> Tuple[Optional[pd.DataFrame], str]:
        """
        Read an Excel file and return a DataFrame with the method used.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Tuple of (DataFrame or None, method_name)
        """
        # Try multiple methods in order of preference
        methods = [
            ("pandas with engine detection", self._read_with_pandas),
            ("xlwings", self._read_with_xlwings)
        ]
        
        for method_name, method_func in methods:
            try:
                df = method_func(file_path)
                if df is not None:
                    return df, method_name
            except Exception as e:
                # Continue to next method
                continue
        
        # If all methods failed, add error
        self.errors.append(f"Error reading {file_path}: All reading methods failed")
        return None, "failed"
    
    def _read_with_pandas(self, file_path: Path) -> Optional[pd.DataFrame]:
        """Read Excel file using pandas with appropriate engine."""
        file_ext = file_path.suffix.lower()
        
        if file_ext == '.xlsx':
            # Use openpyxl for .xlsx files
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_ext == '.xls':
            # Try xlrd first for .xls files, fallback to openpyxl if it fails
            try:
                df = pd.read_excel(file_path, engine='xlrd')
            except Exception:
                # Some .xls files are actually XML-based Excel files
                df = pd.read_excel(file_path, engine='openpyxl')
        else:
            # Try to auto-detect
            df = pd.read_excel(file_path)
        
        # Clean column names
        return self.clean_column_names(df)
    
    def _read_with_xlwings(self, file_path: Path) -> Optional[pd.DataFrame]:
        """Read Excel file using xlwings as fallback (macOS only)."""
        try:
            import xlwings as xw
            
            # Open the workbook
            wb = xw.Book(str(file_path))
            
            # Get the first worksheet
            ws = wb.sheets[0]
            
            # Read all data from the worksheet
            data = ws.used_range.value
            
            # Close the workbook
            wb.close()
            
            if data is None:
                return None
            
            # Convert to DataFrame
            if isinstance(data, list) and len(data) > 0:
                # First row as headers
                headers = data[0]
                rows = data[1:]
                
                # Create DataFrame
                df = pd.DataFrame(rows, columns=headers)
                # Clean column names
                return self.clean_column_names(df)
            else:
                return None
                
        except ImportError:
            # xlwings not available (e.g., in Linux containers)
            raise ImportError("xlwings not available in this environment")
        except Exception as e:
            # xlwings failed, return None to try next method
            raise e
    
    def _read_with_csv(self, file_path: Path) -> Optional[pd.DataFrame]:
        """
        Read CSV file using pandas with automatic detection.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            DataFrame or None if reading fails
        """
        # Common encodings to try
        encodings = ['utf-8', 'latin-1', 'windows-1252', 'utf-16']
        
        # Common delimiters to try
        delimiters = [',', ';', '\t', '|']
        
        # Try different combinations of encoding and delimiter
        for encoding in encodings:
            for delimiter in delimiters:
                try:
                    df = pd.read_csv(
                        file_path,
                        encoding=encoding,
                        delimiter=delimiter,
                        quotechar='"',
                        skipinitialspace=True,
                        on_bad_lines='skip'  # Skip problematic lines
                    )
                    
                    # Check if we got a reasonable result
                    if len(df.columns) > 1 and len(df) > 0:
                        return df
                        
                except Exception:
                    continue
        
        # If all combinations failed, try pandas auto-detection
        try:
            df = pd.read_csv(file_path, encoding='utf-8', on_bad_lines='skip')
            return df
        except Exception:
            # Last resort: try with different quote characters
            try:
                df = pd.read_csv(
                    file_path, 
                    encoding='utf-8', 
                    quotechar="'", 
                    on_bad_lines='skip'
                )
                return df
            except Exception:
                return None
    
    def merge_dataframes(self, dataframes: List[pd.DataFrame]) -> pd.DataFrame:
        """
        Merge multiple DataFrames, handling column mismatches.
        
        Args:
            dataframes: List of DataFrames to merge
            
        Returns:
            Merged DataFrame
        """
        if not dataframes:
            return pd.DataFrame()
            
        if len(dataframes) == 1:
            return dataframes[0]
            
        # Use pandas.concat with sort=False to preserve column order
        # This automatically handles missing columns by filling with NaN
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
        
        return merged_df
    
    def merge_subdirectory(self, subdir_path: Path) -> bool:
        """
        Merge all supported files (Excel and CSV) in a subdirectory.
        
        Args:
            subdir_path: Path to the subdirectory
            
        Returns:
            bool: True if successful, False otherwise
        """
        supported_files = self.get_supported_files(subdir_path)
        
        if not supported_files:
            print(f"  No supported files found in {subdir_path.name}")
            return True
            
        print(f"  Found {len(supported_files)} supported files")
        
        # Check if there's only one file - no merge needed
        if len(supported_files) == 1:
            print(f"  ℹ️  Only one file found in {subdir_path.name} - no merge needed")
            self.skipped_single_file += 1
            return True
        
        # Read all supported files
        dataframes = []
        for file_path in supported_files:
            df, method_used = self.read_file_with_method(file_path)
            if df is not None:
                dataframes.append(df)
                print(f"    ✓ Read {file_path.name} ({len(df)} rows, {len(df.columns)} columns) using {method_used}")
            else:
                print(f"    ✗ Failed to read {file_path.name}")
        
        if not dataframes:
            print(f"  No valid files to merge in {subdir_path.name}")
            return False
            
        # Merge the dataframes
        merged_df = self.merge_dataframes(dataframes)
        
        # Create output filename
        output_filename = f"{subdir_path.name}_merged.xlsx"
        output_path = subdir_path / output_filename
        
        # Save the merged file
        try:
            merged_df.to_excel(output_path, index=False)
            print(f"  ✓ Created merged file: {output_filename} ({len(merged_df)} rows, {len(merged_df.columns)} columns)")
            self.total_files_merged += len(dataframes)
            return True
        except Exception as e:
            self.errors.append(f"Error saving merged file {output_path}: {str(e)}")
            print(f"  ✗ Failed to save merged file: {str(e)}")
            return False
    
    def process_all_subdirectories(self) -> Dict[str, any]:
        """
        Process all subdirectories in the root directory.
        
        Returns:
            Dictionary with processing results
        """
        if not self.validate_directory():
            return {
                'success': False,
                'errors': self.errors,
                'processed_subdirs': 0,
                'total_files_merged': 0
            }
        
        print(f"Processing directory: {self.root_directory}")
        print(f"Found {len([d for d in self.root_directory.iterdir() if d.is_dir()])} subdirectories")
        print("-" * 50)
        
        subdirs = [d for d in self.root_directory.iterdir() if d.is_dir()]
        
        for subdir in subdirs:
            print(f"Processing subdirectory: {subdir.name}")
            success = self.merge_subdirectory(subdir)
            if success:
                self.processed_subdirs += 1
            print()
        
        return {
            'success': True,
            'processed_subdirs': self.processed_subdirs,
            'total_files_merged': self.total_files_merged,
            'skipped_single_file': self.skipped_single_file,
            'errors': self.errors,
            'total_subdirs': len(subdirs)
        }
    
    def print_summary(self, results: Dict[str, any]):
        """Print a summary of the processing results."""
        print("=" * 50)
        print("PROCESSING SUMMARY")
        print("=" * 50)
        
        if results['success']:
            print(f"✓ Successfully processed {results['processed_subdirs']} out of {results['total_subdirs']} subdirectories")
            print(f"✓ Merged {results['total_files_merged']} files")
            if results.get('skipped_single_file', 0) > 0:
                print(f"ℹ️  Skipped {results['skipped_single_file']} subdirectories with only one file (no merge needed)")
        else:
            print("✗ Processing failed")
            
        if results['errors']:
            print(f"\n⚠️  {len(results['errors'])} errors encountered:")
            for error in results['errors']:
                print(f"  - {error}")
        
        print("=" * 50)
