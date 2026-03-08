#!/usr/bin/env python3
"""
List Reference Columns - Show columns of the reference file in each subdirectory.

For each subdirectory under the given root, finds the reference file (the one with
the most columns, same logic as the merger), reads it with the same cleaning, and
prints its columns in an organized way. Use this to choose which columns to put in
columns.yaml for --use-columns.

Usage: python list_reference_columns.py <directory_path> [--yaml]
"""

import sys
import argparse
from pathlib import Path
from typing import List, Tuple, Optional

# Add the src directory to Python path for imports
sys.path.insert(0, str(Path(__file__).parent))

from file_merger import FileMerger


def get_reference_columns_per_subdir(
    merger: FileMerger,
) -> List[Tuple[Path, Path, List[str]]]:
    """
    For each subdirectory, get the reference file (most columns) and its column list.
    Returns list of (subdir_path, reference_file_path, columns).
    """
    subdirs = sorted(merger._get_subdirectories_to_process())
    results: List[Tuple[Path, Path, List[str]]] = []

    for subdir in subdirs:
        supported = merger.get_supported_files(subdir)
        # Skip Excel lock files
        supported = [p for p in supported if not p.name.startswith("~$")]
        if not supported:
            results.append((subdir, Path(""), []))
            continue

        dataframes: List[Tuple[Path, List[str]]] = []
        for file_path in supported:
            df, _ = merger.read_file_with_method(file_path)
            if df is not None:
                dataframes.append((file_path, df.columns.tolist()))

        if not dataframes:
            results.append((subdir, Path(""), []))
            continue

        reference_path, reference_columns = max(
            dataframes, key=lambda x: len(x[1])
        )
        results.append((subdir, reference_path, reference_columns))

    return results


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Print columns of the reference file (most columns) in each subdirectory",
        epilog="""
Examples:
  python list_reference_columns.py ./Data
  python list_reference_columns.py ./Data --yaml
  python list_reference_columns.py ./Data --per-file   # columns for every file (see naming variants)
        """,
    )
    parser.add_argument(
        "directory",
        help="Root directory containing subdirectories with Excel/CSV files",
    )
    parser.add_argument(
        "--yaml",
        action="store_true",
        help="Print a YAML-ready 'columns:' block at the end (union of all reference columns)",
    )
    parser.add_argument(
        "--per-file",
        action="store_true",
        help="For each subdir, list columns for every file (not just the reference). Use to see naming variants for column_aliases.",
    )
    args = parser.parse_args()

    directory_path = Path(args.directory).resolve()
    if not directory_path.exists() or not directory_path.is_dir():
        print(f"Error: Directory does not exist or is not a directory: {directory_path}")
        sys.exit(1)

    merger = FileMerger(
        str(directory_path),
        clean_columns=True,
        validate_structure=False,
    )

    if not merger.validate_directory():
        for err in merger.errors:
            print(f"Error: {err}")
        sys.exit(1)

    if args.per_file:
        # Per-file mode: for each subdir, print every file and its columns
        subdirs = sorted(merger._get_subdirectories_to_process())
        print("Columns per file (all subdirectories)")
        print("=" * 60)
        print(f"Root: {directory_path}\n")
        all_columns_union = set()
        for subdir in subdirs:
            file_columns_list = merger._get_all_file_columns_for_subdir(subdir)
            print(f"Subdirectory: {subdir.name}")
            if not file_columns_list:
                print("  (no supported files or could not read any file)")
            else:
                for file_path, columns in file_columns_list:
                    print(f"  {file_path.name}: {columns}")
                    for c in columns:
                        all_columns_union.add(c)
            print()
        if args.yaml and all_columns_union:
            print("=" * 60)
            print("YAML block (copy to columns.yaml):")
            print("-" * 60)
            print("columns:")
            for col in sorted(all_columns_union):
                print(f"  - {col}")
            print("-" * 60)
    else:
        results = get_reference_columns_per_subdir(merger)
        all_columns_union = set()

        print("Reference columns per subdirectory")
        print("=" * 60)
        print(f"Root: {directory_path}\n")

        for subdir, ref_path, columns in results:
            sep = "=" * 60
            print(f"{sep}")
            print(f"Subdirectory: {subdir.name}")
            if ref_path and ref_path.name:
                print(f"Reference file: {ref_path.name} ({len(columns)} columns)")
                for i, col in enumerate(columns, 1):
                    print(f"  {i:3}. {col}")
                    all_columns_union.add(col)
            else:
                print("  (no supported files or could not read any file)")
            print()

        if args.yaml and all_columns_union:
            print("=" * 60)
            print("YAML block (copy to columns.yaml):")
            print("-" * 60)
            print("columns:")
            for col in sorted(all_columns_union):
                print(f"  - {col}")
            print("-" * 60)


if __name__ == "__main__":
    main()
