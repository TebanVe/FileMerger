#!/usr/bin/env python3
"""
File Merger - Main Script

This script merges multiple Excel and CSV files from subdirectories into consolidated files.
Usage: python merge_excel_files.py <directory_path>
"""

import sys
import argparse
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any

# Add the src directory to Python path for imports
sys.path.insert(0, str(Path(__file__).parent))

from file_merger import FileMerger


def load_columns_from_yaml(path: Path) -> List[str]:
    """
    Load column names from a YAML file with a top-level 'columns' list.
    Raises ValueError if the file is missing or invalid.
    """
    config = load_config_from_yaml(path, require_columns=True)
    return config['columns']


def load_config_from_yaml(path: Path, require_columns: bool = True) -> Dict[str, Any]:
    """
    Load merge config from a YAML file.

    YAML may contain:
      columns: [list of column names] (required if require_columns=True)
      large_file_threshold_mb: optional float
      subdir_include: optional list of subdirectory names to process only
      subdir_exclude: optional list of subdirectory names to skip
      group_by: optional list of column names for aggregation (rows of the pivot)
      max_rows_before_aggregate: optional int (aggregate when merged rows exceed this)
      aggregate_columns: optional dict column_name -> agg_func (value columns and how to aggregate: sum, mean, count, first, last, min, max)
      sort_by: optional list of column names (or single name) to sort merged output by before save
      sort_ascending: optional bool; false = descending (e.g. largest first), true = ascending
      subdirectories: optional dict of subdir_name -> { columns, group_by, max_rows_before_aggregate, aggregate_columns, sort_by, sort_ascending }
      column_aliases: optional dict variant_name -> canonical_name (for normalizing column names)
      column_canonical_patterns: optional list of { base, canonical } (e.g. "Name (suffix)" -> canonical)

    Returns:
        Dict with keys: columns, large_file_threshold_mb, subdir_include, subdir_exclude,
        group_by, max_rows_before_aggregate, aggregate_columns, sort_by, sort_ascending,
        subdirectory_config, column_aliases, column_canonical_patterns
    """
    import yaml
    if not path.exists():
        raise ValueError(f"Config file not found: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    if not data:
        data = {}
    if require_columns and 'columns' not in data:
        raise ValueError(f"YAML file {path} must contain a 'columns' key with a list")
    cols = data.get('columns') or []
    if isinstance(cols, list):
        columns_list = [str(c).strip() for c in cols if c]
    else:
        columns_list = []
    threshold = data.get('large_file_threshold_mb')
    if threshold is not None:
        try:
            threshold = float(threshold)
            if threshold <= 0:
                threshold = None
        except (TypeError, ValueError):
            threshold = None
    subdir_include = data.get('subdir_include')
    if subdir_include is not None and not isinstance(subdir_include, list):
        subdir_include = [str(subdir_include)]
    subdir_exclude = data.get('subdir_exclude')
    if subdir_exclude is not None and not isinstance(subdir_exclude, list):
        subdir_exclude = [str(subdir_exclude)]
    group_by = data.get('group_by')
    if group_by is not None and not isinstance(group_by, list):
        group_by = [str(group_by)] if group_by else None
    max_rows = data.get('max_rows_before_aggregate')
    if max_rows is not None:
        try:
            max_rows = int(max_rows)
            if max_rows < 1:
                max_rows = None
        except (TypeError, ValueError):
            max_rows = None
    aggregate_columns = data.get('aggregate_columns')
    if aggregate_columns is not None and isinstance(aggregate_columns, dict):
        aggregate_columns = {str(k): str(v).lower() for k, v in aggregate_columns.items()}
    else:
        aggregate_columns = {}
    sort_by = data.get('sort_by')
    if sort_by is not None:
        sort_by = sort_by if isinstance(sort_by, list) else [str(sort_by)]
        sort_by = [str(s).strip() for s in sort_by if s]
    else:
        sort_by = []
    sort_ascending = data.get('sort_ascending')
    if sort_ascending is not None:
        sort_ascending = bool(sort_ascending)
    else:
        sort_ascending = False
    subdirectory_config = {}
    for name, opts in (data.get('subdirectories') or {}).items():
        if not isinstance(opts, dict):
            continue
        entry = {}
        if 'columns' in opts and isinstance(opts['columns'], list):
            entry['columns'] = [str(c).strip() for c in opts['columns'] if c]
        if 'group_by' in opts:
            g = opts['group_by']
            entry['group_by'] = g if isinstance(g, list) else [str(g)] if g else None
        if 'max_rows_before_aggregate' in opts:
            try:
                entry['max_rows_before_aggregate'] = int(opts['max_rows_before_aggregate'])
            except (TypeError, ValueError):
                pass
        if 'aggregate_columns' in opts and isinstance(opts['aggregate_columns'], dict):
            entry['aggregate_columns'] = {str(k): str(v).lower() for k, v in opts['aggregate_columns'].items()}
        if 'sort_by' in opts and opts['sort_by'] is not None:
            s = opts['sort_by']
            entry['sort_by'] = s if isinstance(s, list) else [str(s)]
        if 'sort_ascending' in opts and opts['sort_ascending'] is not None:
            entry['sort_ascending'] = bool(opts['sort_ascending'])
        subdirectory_config[str(name)] = entry
    column_aliases = data.get('column_aliases')
    if column_aliases is not None and isinstance(column_aliases, dict):
        column_aliases = {str(k): str(v) for k, v in column_aliases.items()}
    else:
        column_aliases = {}
    column_canonical_patterns = data.get('column_canonical_patterns')
    if column_canonical_patterns is not None and isinstance(column_canonical_patterns, list):
        patterns = []
        for item in column_canonical_patterns:
            if isinstance(item, dict) and 'base' in item and 'canonical' in item:
                patterns.append({'base': str(item['base']), 'canonical': str(item['canonical'])})
        column_canonical_patterns = patterns
    else:
        column_canonical_patterns = []
    return {
        'columns': columns_list,
        'large_file_threshold_mb': threshold,
        'subdir_include': subdir_include,
        'subdir_exclude': subdir_exclude,
        'group_by': group_by,
        'max_rows_before_aggregate': max_rows,
        'aggregate_columns': aggregate_columns,
        'sort_by': sort_by,
        'sort_ascending': sort_ascending,
        'subdirectory_config': subdirectory_config,
        'column_aliases': column_aliases,
        'column_canonical_patterns': column_canonical_patterns,
    }


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
    parser.add_argument(
        '--large-file-threshold',
        type=float,
        default=None,
        metavar='MB',
        help='Use chunked/streaming read for files at or above this size (MB). Overrides YAML if set.'
    )
    parser.add_argument(
        '--subdirs',
        type=str,
        default=None,
        metavar='NAME1,NAME2',
        help='Only process these subdirectories (comma-separated names). Overrides YAML if set.'
    )
    parser.add_argument(
        '--exclude-subdirs',
        type=str,
        default=None,
        metavar='NAME1,NAME2',
        help='Skip these subdirectories (comma-separated names). Overrides YAML if set.'
    )
    parser.add_argument(
        '--group-by',
        type=str,
        default=None,
        metavar='COL1,COL2',
        help='When merged rows exceed --max-rows-before-aggregate, aggregate by these columns (comma-separated).'
    )
    parser.add_argument(
        '--max-rows-before-aggregate',
        type=int,
        default=None,
        metavar='N',
        help='If set with --group-by, aggregate merged result when row count exceeds N.'
    )
    parser.add_argument(
        '--config-file',
        type=str,
        default=None,
        metavar='PATH',
        help='YAML file for subdir/group_by config when not using --use-columns (subdir_include, group_by, etc.).'
    )

    args = parser.parse_args()

    # Convert to Path object and resolve
    directory_path = Path(args.directory).resolve()

    # Load config from YAML when --use-columns or --config-file
    config = {}
    if args.use_columns:
        columns_path = Path(args.columns_file).resolve() if args.columns_file else Path.cwd() / 'columns.yaml'
        try:
            config = load_config_from_yaml(columns_path, require_columns=True)
            if not config.get('columns'):
                print(f"Error: columns file {columns_path} has an empty 'columns' list.")
                sys.exit(2)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(2)
    elif args.config_file:
        config_path = Path(args.config_file).resolve()
        try:
            config = load_config_from_yaml(config_path, require_columns=False)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(2)

    columns_to_export = config.get('columns') if args.use_columns else None
    large_file_threshold_mb = args.large_file_threshold or config.get('large_file_threshold_mb')
    subdir_include = None
    if args.subdirs:
        subdir_include = [s.strip() for s in args.subdirs.split(',') if s.strip()]
    else:
        subdir_include = config.get('subdir_include')
    subdir_exclude = None
    if args.exclude_subdirs:
        subdir_exclude = [s.strip() for s in args.exclude_subdirs.split(',') if s.strip()]
    else:
        subdir_exclude = config.get('subdir_exclude')
    group_by = None
    if args.group_by:
        group_by = [s.strip() for s in args.group_by.split(',') if s.strip()]
    else:
        group_by = config.get('group_by')
    max_rows_before_aggregate = args.max_rows_before_aggregate if args.max_rows_before_aggregate is not None else config.get('max_rows_before_aggregate')
    aggregate_columns = config.get('aggregate_columns') or {}
    sort_by = config.get('sort_by') or []
    sort_ascending = config.get('sort_ascending', False)
    subdirectory_config = config.get('subdirectory_config') or {}
    column_aliases = config.get('column_aliases') or {}
    column_canonical_patterns = config.get('column_canonical_patterns') or []

    print("File Merger (Excel & CSV)")
    print("=" * 50)
    print(f"Target directory: {directory_path}")
    if columns_to_export:
        print(f"Columns filter: {len(columns_to_export)} columns from YAML")
    if large_file_threshold_mb is not None:
        print(f"Large-file threshold: {large_file_threshold_mb} MB (chunked/streaming read)")
    if subdir_include:
        print(f"Subdirs (only): {subdir_include}")
    if subdir_exclude:
        print(f"Subdirs (exclude): {subdir_exclude}")
    if group_by and max_rows_before_aggregate is not None:
        pivot_desc = f"group_by={group_by}, when rows > {max_rows_before_aggregate}"
        if aggregate_columns:
            pivot_desc += f", value columns/agg={aggregate_columns}"
        print(f"Pivot/aggregate: {pivot_desc}")
    if sort_by:
        print(f"Sort: by={sort_by}, ascending={sort_ascending}")
    if subdirectory_config:
        print(f"Per-subdir config: {list(subdirectory_config.keys())}")
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
        columns_to_export=columns_to_export,
        large_file_threshold_mb=large_file_threshold_mb,
        subdir_include=subdir_include,
        subdir_exclude=subdir_exclude,
        group_by=group_by,
        max_rows_before_aggregate=max_rows_before_aggregate,
        aggregate_columns=aggregate_columns,
        subdirectory_config=subdirectory_config,
        sort_by=sort_by,
        sort_ascending=sort_ascending,
        column_aliases=column_aliases,
        column_canonical_patterns=column_canonical_patterns
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
