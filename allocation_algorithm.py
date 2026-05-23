
"""
Donation allocation report for "Tools for School (3).xlsx".

This script reads the Secondary and Primary sheets, cleans messy inventory data,
and prints a transparent rules-based allocation report for:

1. Bali Children Foundation
2. Hope Children's Home

The goal is to stay close to the original "To which organization" column while
making the allocation rules clear and repeatable.

Run:
    python allocation_report.py

Or pass a workbook path:
    python allocation_report.py "/path/to/Tools for School (3).xlsx"
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from typing import Any

import pandas as pd


DEFAULT_WORKBOOK = Path("/Users/aidenpangu/Downloads/Tools for School (3).xlsx")

BALI = "Bali Children Foundation"
HOPE = "Hope Children's Home"
BOTH = "Both"
REVIEW = "Review Needed"

# Hope minimum essential coverage targets requested by the charity drive rules.
HOPE_MINIMUMS = {
    "Notebooks": 80,
    "Pencils/Pens": 50,
    "Erasers": 48,
}

# Higher scores mean the item is more essential for basic school readiness.
PRIORITY_SCORES = {
    "Notebooks": 5,
    "Pencils/Pens": 5,
    "Erasers": 4,
    "Sharpeners": 4,
    "Rulers": 3,
    "Geometry tools / protractors": 3,
    "Pencil cases": 3,
    "Stationery sets": 3,
    "Children's English books": 4,
    "Coloring supplies": 3,
    "Art supplies": 2,
    "DT tools": 2,
    "Science toys / educational toys": 2,
    "General toys": 1,
    "Unknown": 0,
}

ESSENTIAL_CATEGORIES = {
    "Notebooks",
    "Pencils/Pens",
    "Erasers",
    "Sharpeners",
    "Rulers",
    "Geometry tools / protractors",
    "Pencil cases",
    "Stationery sets",
}

# These categories always go to Bali Children Foundation according to the rules.
FORCED_BCF_CATEGORIES = {
    "Children's English books",
    "Coloring supplies",
    "Art supplies",
    "DT tools",
    "Science toys / educational toys",
    "General toys",
}


def normalize_text(value: Any) -> str:
    """Turn messy spreadsheet text into lowercase text that is easier to match."""
    if value is None:
        return ""

    # Pandas stores blank Excel cells as NaN, and NaN is not equal to itself.
    if isinstance(value, float) and math.isnan(value):
        return ""

    text = str(value).strip().lower()

    # Normalize curly apostrophes and collapse repeated spaces.
    text = text.replace("'", "'").replace("'", "'")
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_destination(value: Any) -> str:
    """Normalize organization names from the workbook into a small set of labels."""
    text = normalize_text(value)

    if not text:
        return ""

    if text == "both" or "both" in text:
        return BOTH

    # The workbook uses "Hope Children's Orphanage"; the requested report uses
    # "Hope Children's Home".
    if "hope" in text:
        return HOPE

    # Catch "Bali Child", "Bali children foundation", and similar wording.
    if "bali" in text or "bcf" in text:
        return BALI

    return REVIEW


def clean_item_category(raw_item: Any, notes: Any = "") -> str:
    """
    Convert messy item names into standard categories.

    The item name is the main signal. Notes are included as a backup because
    some rows describe sets or enrichment items more clearly in the Notes cell.
    """
    item = normalize_text(raw_item)
    note = normalize_text(notes)
    combined = f"{item} {note}".strip()

    if not combined:
        return "Unknown"

    # The raw item name is the main signal. Notes can contain many supplies in a
    # set, so using notes too early can accidentally turn "Pencils/Pens" into
    # "Pencil cases" just because the note mentions a case.
    if any(word in item for word in ["coloring pencil", "colouring pencil", "color pencil", "colour pencil"]):
        return "Coloring supplies"

    if any(word in item for word in ["children's book", "childrens book"]):
        return "Children's English books"

    if "notebook" in item or "exercise book" in item or "excersise book" in item:
        return "Notebooks"

    if "eraser" in item:
        return "Erasers"

    if "sharpener" in item:
        return "Sharpeners"

    if "ruler" in item:
        return "Rulers"

    if any(word in item for word in ["protractor", "potractor", "geometry"]):
        return "Geometry tools / protractors"

    if "pencil case" in item:
        return "Pencil cases"

    if "stationary set" in item or "stationery set" in item:
        return "Stationery sets"

    if "dt tool" in item:
        return "DT tools"

    if any(word in item for word in ["glitter", "paint", "sticker", "popsicle", "frame", "tooth pick", "toothpick", "confetti"]):
        return "Art supplies"

    if any(word in item for word in ["science", "tech"]):
        return "Science toys / educational toys"

    if "puzzle" in item:
        return "Science toys / educational toys"

    if "toy" in item:
        return "General toys"

    if any(word in item for word in ["pencils/pens", "pencil/pen", "pencil", "pen", "marker", "highlighter"]):
        return "Pencils/Pens"

    # If the raw item is vague or unknown, use notes as a backup.
    if any(word in combined for word in ["science", "tech", "magnet physics", "math blocks", "abacus", "magic kit"]):
        return "Science toys / educational toys"

    if any(word in combined for word in ["educational toy", "learning toy"]):
        return "Science toys / educational toys"

    if any(word in combined for word in ["toy", "slime", "cars", "card shuffler", "maracas", "tent"]):
        return "General toys"

    if any(word in combined for word in ["coloring pencil", "colouring pencil", "color pencil", "colour pencil"]):
        return "Coloring supplies"

    if any(word in combined for word in ["glitter", "paint", "sticker", "popsicle", "frame", "tooth pick", "toothpick", "confetti"]):
        return "Art supplies"

    if any(word in combined for word in ["children's book", "childrens book", "abc alphabet book", "reading book", "lego book"]):
        return "Children's English books"

    if "puzzle" in combined:
        return "Science toys / educational toys"

    if any(word in combined for word in ["pencils/pens", "pencil/pen", "pencil", "pen", "marker", "highlighter"]):
        # Coloring pencils were already caught above, so remaining pencils/pens
        # count as writing supplies.
        return "Pencils/Pens"

    return "Unknown"


def _extract_destination_amounts(notes: str) -> dict[str, float]:
    """
    Find simple quantity-and-destination phrases inside notes.

    This helper supports both phrase orders:
    - "3 to Bali ... 2 to Hope"
    - "Hope gets 5 ... Bali gets 15"
    """
    amounts = {BALI: 0.0, HOPE: 0.0}

    # Examples: "3 to Bali Children Foundation", "10 packs to Hope"
    number_to_dest = re.finditer(
        r"(?P<num>\d+(?:\.\d+)?)\s*(?:boxes?|packs?|sets?|pieces?)?\s*to\s*(?P<dest>bali child(?:ren)? foundation|bali child|bali|hope children's orphanage|hope orphanage|hope)",
        notes,
    )
    for match in number_to_dest:
        dest = normalize_destination(match.group("dest"))
        if dest in amounts:
            amounts[dest] += float(match.group("num"))

    # Examples: "Hope gets 5", "Bali children foundation gets 15"
    dest_gets_number = re.finditer(
        r"(?P<dest>bali child(?:ren)? foundation|bali child|bali|hope children's orphanage|hope orphanage|hope)\s*gets?\s*(?P<num>\d+(?:\.\d+)?)",
        notes,
    )
    for match in dest_gets_number:
        dest = normalize_destination(match.group("dest"))
        if dest in amounts:
            amounts[dest] += float(match.group("num"))

    return amounts


def extract_manual_split(notes: Any, quantity: Any) -> dict[str, Any] | None:
    """
    Detect clear manual splits from the Notes column.

    Returns None when no clear split is found. When a split is found, the values
    are scaled to the row's Quantity. If the note amounts already add up to the
    row quantity, no scaling is needed. If the note amounts are box/pack counts,
    the same ratio is used to divide the row quantity.
    """
    text = normalize_text(notes)
    qty = pd.to_numeric(quantity, errors="coerce")

    if not text or pd.isna(qty) or qty <= 0:
        return None

    amounts = _extract_destination_amounts(text)
    split_total = amounts[BALI] + amounts[HOPE]

    # A split must mention both organizations with positive amounts.
    if amounts[BALI] <= 0 or amounts[HOPE] <= 0 or split_total <= 0:
        return None

    hope_qty = qty * amounts[HOPE] / split_total
    bcf_qty = qty * amounts[BALI] / split_total

    return {
        "algorithm_destination": BOTH,
        "hope_quantity": hope_qty,
        "bcf_quantity": bcf_qty,
        "reason": (
            "Manual split preserved from Notes "
            f"({amounts[HOPE]:g} Hope / {amounts[BALI]:g} BCF share)."
        ),
    }


def allocate_row(row: pd.Series, hope_progress: dict[str, float]) -> dict[str, Any]:
    """
    Decide the algorithm allocation for one row.

    The algorithm is intentionally transparent:
    1. Forced-Bali categories go to Bali.
    2. Clear manual splits in Notes are preserved.
    3. Essential items keep the original allocation if it is recognizable.
    4. If the original allocation is unclear, Hope gets essentials until its
       minimum coverage target is met.
    5. Anything still unclear is marked Review Needed.
    """
    quantity = pd.to_numeric(row.get("Quantity"), errors="coerce")
    if pd.isna(quantity):
        quantity = 0.0

    category = row["Clean Category"]
    original_destination = row["Original Destination"]
    priority = PRIORITY_SCORES.get(category, 0)

    # Forced destination rules are strongest.
    if category in FORCED_BCF_CATEGORIES:
        return {
            "Algorithm Destination": BALI,
            "Hope Quantity": 0.0,
            "BCF Quantity": quantity,
            "Priority Score": priority,
            "Reason": f"{category} is forced to Bali Children Foundation.",
        }

    # Preserve clear manual splits for non-forced rows.
    manual_split = extract_manual_split(row.get("Notes"), quantity)
    if manual_split:
        hope_progress[category] = hope_progress.get(category, 0.0) + manual_split["hope_quantity"]
        return {
            "Algorithm Destination": manual_split["algorithm_destination"],
            "Hope Quantity": manual_split["hope_quantity"],
            "BCF Quantity": manual_split["bcf_quantity"],
            "Priority Score": priority,
            "Reason": manual_split["reason"],
        }

    # Essential rows should closely replicate the original destination when the
    # original is recognizable.
    if category in ESSENTIAL_CATEGORIES and original_destination in {BALI, HOPE, BOTH}:
        hope_qty = quantity if original_destination == HOPE else 0.0
        bcf_qty = quantity if original_destination == BALI else 0.0

        if original_destination == BOTH:
            # "Both" without a clear split needs review because the script
            # should not invent a row split unless Notes clearly explains it.
            return {
                "Algorithm Destination": REVIEW,
                "Hope Quantity": 0.0,
                "BCF Quantity": 0.0,
                "Priority Score": priority,
                "Reason": "Original says Both, but Notes do not show a clear split.",
            }

        if original_destination == HOPE:
            hope_progress[category] = hope_progress.get(category, 0.0) + hope_qty

        return {
            "Algorithm Destination": original_destination,
            "Hope Quantity": hope_qty,
            "BCF Quantity": bcf_qty,
            "Priority Score": priority,
            "Reason": "Essential item; recognizable original allocation preserved.",
        }

    # If the original is unclear, use Hope's minimum coverage targets for the
    # most important essentials only.
    if category in HOPE_MINIMUMS:
        current = hope_progress.get(category, 0.0)
        target = HOPE_MINIMUMS[category]
        if current < target:
            hope_progress[category] = current + quantity
            return {
                "Algorithm Destination": HOPE,
                "Hope Quantity": quantity,
                "BCF Quantity": 0.0,
                "Priority Score": priority,
                "Reason": f"Original unclear; assigned to Hope to support {category} minimum coverage.",
            }

    # Recognized non-essential rows can still keep a clear original allocation.
    if original_destination in {BALI, HOPE}:
        return {
            "Algorithm Destination": original_destination,
            "Hope Quantity": quantity if original_destination == HOPE else 0.0,
            "BCF Quantity": quantity if original_destination == BALI else 0.0,
            "Priority Score": priority,
            "Reason": "Recognizable original allocation preserved.",
        }

    return {
        "Algorithm Destination": REVIEW,
        "Hope Quantity": 0.0,
        "BCF Quantity": 0.0,
        "Priority Score": priority,
        "Reason": "Unknown or unclear item/destination; human review needed.",
    }


def load_inventory(workbook_path: Path) -> pd.DataFrame:
    """Read both workbook sheets, standardize columns, drop blanks, and combine."""
    sheet_settings = {
        "Secondary": 0,  # headers are in row 1
        "Primary": 1,  # row 1 is blank; headers are in row 2
    }
    frames = []

    for sheet_name, header_row in sheet_settings.items():
        df = pd.read_excel(
            workbook_path,
            sheet_name=sheet_name,
            header=header_row,
            engine="openpyxl",
        )

        # Standardize the one column name that differs between the two sheets.
        df = df.rename(
            columns={
                "In what container": "Container",
                "In what container?": "Container",
                "What?": "Raw Item",
                "To which organization": "Original Destination Raw",
            }
        )

        df["Source Sheet"] = sheet_name

        wanted_columns = [
            "Source Sheet",
            "Container",
            "Raw Item",
            "Quantity",
            "Condition",
            "Original Destination Raw",
            "Notes",
        ]

        # Keep only the columns used by the report. If a column is missing, make
        # it blank so the rest of the script can still run and show Review Needed.
        for column in wanted_columns:
            if column not in df.columns:
                df[column] = pd.NA

        df = df[wanted_columns]

        # Drop rows that are fully blank across the useful inventory fields.
        df = df.dropna(
            how="all",
            subset=["Container", "Raw Item", "Quantity", "Condition", "Original Destination Raw", "Notes"],
        )

        # Some workbooks have trailing rows where only a formula/default value
        # remains in one column. For an inventory report, a real row needs at
        # least an item name or a quantity.
        has_item = df["Raw Item"].apply(normalize_text) != ""
        has_quantity = pd.to_numeric(df["Quantity"], errors="coerce").notna()
        df = df.loc[has_item | has_quantity].copy()

        frames.append(df)

    return pd.concat(frames, ignore_index=True)


def process_inventory(inventory: pd.DataFrame) -> pd.DataFrame:
    """Clean categories/destinations, allocate rows, and calculate match status."""
    df = inventory.copy()

    df["Clean Category"] = df.apply(
        lambda row: clean_item_category(row.get("Raw Item"), row.get("Notes")),
        axis=1,
    )
    df["Original Destination"] = df["Original Destination Raw"].apply(normalize_destination)

    allocations = []
    hope_progress: dict[str, float] = {}

    for _, row in df.iterrows():
        allocations.append(allocate_row(row, hope_progress))

    allocation_df = pd.DataFrame(allocations)
    df = pd.concat([df.reset_index(drop=True), allocation_df], axis=1)

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    # The algorithm matches the original when the normalized destination labels
    # are equal. Review Needed is counted as not matching because it needs a
    # human decision.
    df["Match Original"] = (
        (df["Algorithm Destination"] == df["Original Destination"])
        & (df["Algorithm Destination"] != REVIEW)
    )

    output_columns = [
        "Source Sheet",
        "Container",
        "Raw Item",
        "Clean Category",
        "Quantity",
        "Condition",
        "Original Destination",
        "Algorithm Destination",
        "Hope Quantity",
        "BCF Quantity",
        "Priority Score",
        "Reason",
        "Match Original",
        "Notes",
    ]

    return df[output_columns]


def _print_table(title: str, df: pd.DataFrame) -> None:
    """Print a readable pandas table with a title."""
    print("\n" + "=" * 100)
    print(title)
    print("=" * 100)

    if df.empty:
        print("No rows to show.")
        return

    print(df.to_string(index=False))


def summarize_results(result: pd.DataFrame) -> None:
    """Print all required report sections."""
    pd.set_option("display.max_rows", None)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 80)

    full_table_columns = [
        "Raw Item",
        "Clean Category",
        "Quantity",
        "Condition",
        "Original Destination",
        "Hope Quantity",
        "BCF Quantity",
        "Algorithm Destination",
    ]
    _print_table("1. Full allocation table", result[full_table_columns])

    mismatches = result.loc[~result["Match Original"]].copy()
    _print_table("2. Mismatch review table", mismatches)

    match_percentage = result["Match Original"].mean() * 100 if len(result) else 0
    print("\n" + "=" * 100)
    print("3. Match percentage")
    print("=" * 100)
    print(f"{match_percentage:.1f}% ({result['Match Original'].sum()} of {len(result)} rows)")

    hope_rows = []
    for category, target in HOPE_MINIMUMS.items():
        actual = result.loc[result["Clean Category"] == category, "Hope Quantity"].sum()
        hope_rows.append(
            {
                "Clean Category": category,
                "Hope Minimum Target": target,
                "Hope Quantity": actual,
                "Meets Minimum?": actual >= target,
                "Difference": actual - target,
            }
        )

    _print_table("4. Hope minimum coverage check", pd.DataFrame(hope_rows))


def main() -> None:
    """Command-line entry point."""
    parser = argparse.ArgumentParser(description="Generate a school supply donation allocation report.")
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

    inventory = load_inventory(workbook_path)
    result = process_inventory(inventory)
    summarize_results(result)


if __name__ == "__main__":
    main()



# if i want to run it     *    python allocation_report.py "Tools for School (3).xlsx"        *
