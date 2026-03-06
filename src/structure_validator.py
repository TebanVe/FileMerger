"""
Structure Validator - Column structure validation before merge.

Validates that all files in a subdirectory have the same set of column names
(after cleaning). Reports file path, column name, Excel column reference, and
optional plural/singular hints.
"""

from pathlib import Path
from typing import List, Optional, Dict, Tuple
from dataclasses import dataclass
from typing import Literal

import pandas as pd


def index_to_excel_column(index: int) -> str:
    """
    Convert 0-based column index to Excel column letter(s).
    0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, etc.
    """
    result = []
    n = index
    while True:
        result.append(chr(ord('A') + (n % 26)))
        n = n // 26
        if n == 0:
            break
        n -= 1
    return ''.join(reversed(result))


def _is_plural_singular_pair(a: str, b: str) -> bool:
    """
    Return True if one name is the other plus/minus trailing 's' or 'es'.
    Used only for hint messages in validation reports.
    """
    if not a or not b:
        return False
    a, b = a.strip(), b.strip()
    if a == b:
        return False
    if a == b + 's' or a + 's' == b:
        return True
    if a == b + 'es' or a + 'es' == b:
        return True
    if len(b) > 1 and b.endswith('s') and a == b[:-1]:
        return True
    if len(a) > 1 and a.endswith('s') and b == a[:-1]:
        return True
    return False


def file_has_plural_singular_conflict(columns: List[str]) -> Optional[Tuple[str, str]]:
    """
    If this file has two columns that are plural/singular of each other (e.g. Brand and Brands),
    return that pair so the user can be asked which to keep. Otherwise return None.
    """
    cols = [c for c in columns if c and str(c).strip()]
    for i, a in enumerate(cols):
        for b in cols[i + 1:]:
            if _is_plural_singular_pair(str(a).strip(), str(b).strip()):
                return (str(a).strip(), str(b).strip())
    return None


def _plural_singular_groups(all_columns: List[str]) -> List[List[str]]:
    """Partition column names into groups where names in a group are plural/singular of each other."""
    all_columns = [str(c).strip() for c in all_columns if c and str(c).strip()]
    groups: List[List[str]] = []
    for name in all_columns:
        found = False
        for g in groups:
            if any(_is_plural_singular_pair(name, x) for x in g):
                g.append(name)
                found = True
                break
        if not found:
            groups.append([name])
    return groups


def compute_canonical_plural_singular_renames(
    file_columns_list: List[List[str]],
) -> Dict[str, str]:
    """
    For each set of column names that are plural/singular variants across files,
    choose the canonical name (the one that appears in the most files). Return
    a mapping from each non-canonical name to the canonical one.
    """
    all_names = set()
    for cols in file_columns_list:
        all_names.update(str(c).strip() for c in cols if c and str(c).strip())

    renames: Dict[str, str] = {}
    groups = _plural_singular_groups(list(all_names))

    for group in groups:
        if len(group) < 2:
            continue
        # Count how many files contain each name in this group
        count_per_name: Dict[str, int] = {n: 0 for n in group}
        for cols in file_columns_list:
            col_set = {str(c).strip() for c in cols if c}
            for n in group:
                if n in col_set:
                    count_per_name[n] += 1
        # Canonical = name with maximum count (majority)
        canonical = max(group, key=lambda n: count_per_name[n])
        for n in group:
            if n != canonical:
                renames[n] = canonical
    return renames


@dataclass
class StructureIssue:
    """A single column structure mismatch for one file."""
    file_path: Path
    column_name: str
    column_index: int
    excel_ref: str
    issue_type: Literal["missing", "extra"]
    message: str
    hint: Optional[str] = None


@dataclass
class StructureValidationResult:
    """Result of comparing column structure across files."""
    success: bool
    reference_file: Path
    reference_columns: List[str]
    issues: List[StructureIssue]


def validate_column_structure(
    file_paths: List[Path],
    dataframes: List[pd.DataFrame],
    allow_missing_columns: bool = False,
) -> StructureValidationResult:
    """
    Validate column structure across DataFrames.
    Uses the first file as reference (caller must ensure it's the one with most columns if desired).
    When allow_missing_columns=True, missing columns (in reference but not in file) are allowed
    and do not cause failure; those cells will be left blank on merge.
    """
    if not file_paths or not dataframes or len(file_paths) != len(dataframes):
        return StructureValidationResult(
            success=True,
            reference_file=Path('.'),
            reference_columns=[],
            issues=[]
        )

    reference_path = file_paths[0]
    reference_df = dataframes[0]
    reference_columns = reference_df.columns.tolist()
    reference_set = set(reference_columns)
    issues: List[StructureIssue] = []

    for i in range(1, len(file_paths)):
        path = file_paths[i]
        df = dataframes[i]
        file_columns = df.columns.tolist()
        file_set = set(file_columns)

        missing = reference_set - file_set
        extra = file_set - reference_set

        # Missing columns: only report/fail if not allowed (when allowed, merge leaves them blank)
        if not allow_missing_columns:
            for col in missing:
                idx = reference_columns.index(col) if col in reference_columns else 0
                excel_ref = index_to_excel_column(idx)
                hint = None
                for ex in extra:
                    if _is_plural_singular_pair(col, ex):
                        hint = f"Possible plural/singular: '{col}' vs '{ex}'"
                        break
                issues.append(StructureIssue(
                    file_path=path,
                    column_name=col,
                    column_index=idx,
                    excel_ref=excel_ref,
                    issue_type="missing",
                    message=f"Missing column (expected by reference): {col}",
                    hint=hint
                ))

        for col in extra:
            idx = file_columns.index(col) if col in file_columns else 0
            excel_ref = index_to_excel_column(idx)
            hint = None
            for ref_col in reference_columns:
                if _is_plural_singular_pair(col, ref_col):
                    hint = f"Possible plural/singular: '{col}' vs '{ref_col}'"
                    break
            issues.append(StructureIssue(
                file_path=path,
                column_name=col,
                column_index=idx,
                excel_ref=excel_ref,
                issue_type="extra",
                message=f"Extra column (not in reference): {col}",
                hint=hint
            ))

    # When allow_missing_columns=True, we never block merge (missing columns left blank)
    success = len(issues) == 0 or allow_missing_columns
    return StructureValidationResult(
        success=success,
        reference_file=reference_path,
        reference_columns=reference_columns,
        issues=issues
    )


def format_validation_report(result: StructureValidationResult) -> str:
    """
    Format validation result as a readable report for console/output.
    """
    if result.success:
        return "Column structure: OK (all files match)."

    lines = [
        "COLUMN STRUCTURE MISMATCH",
        "=" * 50,
        f"Reference file: {result.reference_file}",
        f"Reference columns ({len(result.reference_columns)}): " + ", ".join(result.reference_columns),
        "",
        "Issues by file:",
        "-" * 50
    ]

    current_file = None
    for issue in result.issues:
        if issue.file_path != current_file:
            current_file = issue.file_path
            lines.append(f"  File: {issue.file_path}")
        lines.append(f"    Column {issue.excel_ref} ({issue.column_name}): {issue.issue_type} - {issue.message}")
        if issue.hint:
            lines.append(f"      Hint: {issue.hint}")
    lines.append("=" * 50)
    return "\n".join(lines)
