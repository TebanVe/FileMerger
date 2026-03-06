"""
File Merger Package

A Python package for merging Excel and CSV files from subdirectories into consolidated files.
"""

from .file_merger import FileMerger
from .structure_validator import (
    validate_column_structure,
    format_validation_report,
    StructureValidationResult,
    StructureIssue,
    index_to_excel_column,
    file_has_plural_singular_conflict,
    compute_canonical_plural_singular_renames,
)

__version__ = "1.0.0"
__author__ = "FileMerger Project"
__all__ = [
    "FileMerger",
    "validate_column_structure",
    "format_validation_report",
    "StructureValidationResult",
    "StructureIssue",
    "index_to_excel_column",
    "file_has_plural_singular_conflict",
    "compute_canonical_plural_singular_renames",
]
