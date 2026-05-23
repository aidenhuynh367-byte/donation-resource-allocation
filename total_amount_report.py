"""
Total amount report for "Tools for School (3).xlsx".

This script prints the total quantity for each raw item in the workbook.
For example, it answers questions like:

- How many notebooks are there in total?
- How many Pencils/Pens are there in total?
- How many Children's books are there in total?

Run:
    python total_amount_report.py

Or pass a workbook path:
    python total_amount_report.py "/path/to/Tools for School (3).xlsx"
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from allocation_report import DEFAULT_WORKBOOK, load_inventory, normalize_text


def clean_raw_item_for_grouping(raw_item: object) -> str:
    """
    Clean raw item names just enough for grouping totals.

    This does not convert items into broad categories. It only fixes small
    spreadsheet differences like extra spaces and different capitalization.
    """
    text = normalize_text(raw_item)

    if not text:
        return "Unknown"

    # Use a few friendly display names for common rows in the workbook.
    display_names = {
        "children's books": "Children's books",
        "pencils/pens": "Pencils/Pens",
        "erasers": "Erasers",
        "rulers": "Rulers",
        "notebooks": "Notebooks",
        "sharpeners": "Sharpeners",
        "coloring pencil (each pen)": "Coloring pencil (each pen)",
        "pencil case": "Pencil case",
        "toys": "Toys",
        "glitter": "Glitter",
        "potractors set": "Potractors set",
        "stationary sets": "Stationary sets",
        "dt tools": "DT tools",
        "paint supplies": "Paint supplies",
        "science/tech equipment": "Science/tech equipment",
        "puzzles": "Puzzles",
        "school excersise books": "School excersise books",
    }

    return display_names.get(text, text.title())


def build_total_amount_report(workbook_path: Path) -> pd.DataFrame:
    """Read the inventory and calculate total quantity by raw item."""
    inventory = load_inventory(workbook_path)

    inventory["Raw Item Cleaned"] = inventory["Raw Item"].apply(clean_raw_item_for_grouping)
    inventory["Quantity"] = pd.to_numeric(inventory["Quantity"], errors="coerce").fillna(0)

    report = (
        inventory.groupby("Raw Item Cleaned", dropna=False)
        .agg(
            Total_Quantity=("Quantity", "sum"),
            Number_of_Rows=("Raw Item Cleaned", "size"),
        )
        .reset_index()
        .rename(columns={"Raw Item Cleaned": "Raw Item"})
        .sort_values("Raw Item")
    )

    return report


def print_total_amount_report(report: pd.DataFrame) -> None:
    """Print the report in a readable table."""
    pd.set_option("display.max_rows", None)
    pd.set_option("display.width", 120)

    print("\n" + "=" * 70)
    print("Total amount report by raw item")
    print("=" * 70)
    print(report.to_string(index=False))


def main() -> None:
    """Command-line entry point."""
    parser = argparse.ArgumentParser(description="Print total quantity by raw item.")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help=f"Path to the Excel workbook. Default: {DEFAULT_WORKBOOK}",
    )
    args = parser.parse_args()

    workbook_path = Path(args.workbook).expanduser()

    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    report = build_total_amount_report(workbook_path)
    print_total_amount_report(report)


if __name__ == "__main__":
    main()


# IF WANT TO RUN *              python3 total_amount_report.py "Tools for School (3).xlsx"        *

