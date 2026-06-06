"""
Create PM Task List from csv files in a provided path
"""

from enum import Enum
import json
import logging
from pathlib import Path
import csv
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

#-----------------------------------------------------
# Constants & Configurations
#-----------------------------------------------------

# Columns to keep from raw data
KEEP_COLUMNS = ["Task ID", "Task Name", "Assigned To", "Cell"]

# Frequently used file paths
CONFIG_DIR = Path(__file__).parent / "config"
TASK_DICT_PATH = CONFIG_DIR / "task_owners.csv"
WEEKLY_EXPORTS = Path(__file__).parent / "weekly_exports"

# Formatting and styles for writing Excel file
HEADER_FONT = Font(bold=True, size=14, name="Arial Narrow", color="FFFFFF")
NOTE_FONT = Font(bold=True, size=12, name="Arial Narrow", color="FF0000")
TASK_FONT = Font(size=11)
CHECK_FONT = Font(bold=True, size=11)

HEADER_FILL = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
NOTE_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
CHECK_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

CENTER = Alignment(horizontal="center", vertical="bottom")
LEFT = Alignment(horizontal="left", vertical="bottom")
WRAP = Alignment(horizontal="left", vertical="bottom", wrap_text=True)

COLUMN_WIDTHS = {"A": 12.85, "B": 97.28, "C": 21.00, "D": 20.85}
CENTER_COLUMNS = {0, 2, 3}

HEADER_ROW_HEIGHT = 22
NOTE_ROW_HEIGHT = 18
TASK_ROW_HEIGHT = 15

#-----------------------------------------------------
# Enums and Custom Types
#-----------------------------------------------------

# Enum used to verify loaded taskDict has no errors at runtime
class TaskOwner(Enum):
    NIGHTS = "Night Shift Line Tech"
    DAYS = "Day Shift Line Tech"
    IMM = "IMM PM Tech"
    PMTEAM = "PM Team Tech"
    SUPPORT = "PM Support"
    CLEAN = "Cleaning Crew"
    VSPL = "VSPL"
    CAL = "Calibrations"
    MCKEN = "McKendrees"
    MILLER = "Miller Elec."
    RELIABILITY = "Reliability Tech"
    SME = "SME"
    SHUTDOWN = "Shutdown Item"

# List used to define task order on final sheet
CUSTOM_ORDER = [
    "",
    TaskOwner.NIGHTS.value,
    TaskOwner.PMTEAM.value,
    TaskOwner.DAYS.value,
    TaskOwner.VSPL.value,
    TaskOwner.SME.value,
    TaskOwner.SUPPORT.value,
    TaskOwner.CAL.value,
    TaskOwner.MILLER.value,
    TaskOwner.MCKEN.value,
    TaskOwner.RELIABILITY.value,
    TaskOwner.IMM.value,
    TaskOwner.CLEAN.value,
    TaskOwner.SHUTDOWN.value
]

#-----------------------------------------------------
# Logger Setup
#-----------------------------------------------------

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

#-----------------------------------------------------
# Data Loading
#-----------------------------------------------------

def load_task_dict(filepath: Path) -> dict[str, str]:
    """Load and validate task-to-owner mapping from csv file"""
    valid_owners = {member.value for member in TaskOwner}
    task_dict = {}

    with open(filepath, newline="", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            task_id = row["task_id"].strip()
            owner = row["owner"].strip()
            task_dict[task_id] = owner

    # CSV entry validation
    invalid = {k: v for k, v in task_dict.items() if v not in valid_owners}
    if invalid:
        raise ValueError(f"Uknown task owners in {filepath}: {invalid}")
    logger.info("Loaded %d task mappings from %s", len(task_dict), filepath.name)
    return task_dict

def load_csvs(directory: Path) -> pd.DataFrame:
    """Read and concatenate all CSVs in *directory*."""
    csv_files = list(directory.glob("*.csv"))
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {directory}")

    logger.info("Found %d CSV files", len(csv_files))
    return pd.concat(
        (pd.read_csv(f, header=0) for f in csv_files),
        ignore_index=True
    )

def load_section_config(filepath: Path) -> list[dict]:
    """Load section configuration JSON."""
    with open(filepath, "r", encoding="utf-8") as file:
        data = json.load(file)
    if not isinstance(data, list):
        raise ValueError("section_config.json must be a list of section objects")
    return data

#-----------------------------------------------------
# Data Transformation
#-----------------------------------------------------

def assign_owners(df: pd.DataFrame, task_dict: dict[str, str]) -> pd.DataFrame:
    """Populate the Assigned To column from the task lookup file"""
    df = df.copy()
    df["Assigned To"] = df["Task ID"].map(task_dict)
    unmapped = df["Assigned To"].isna().sum()
    if unmapped:
        logger.warning("%d tasks had no owner mapping", unmapped)
    return df

def sort_by_owner(df: pd.DataFrame, order: list[str]) -> pd.DataFrame:
    """Sort rows by custom owner order; unmapped owners go to the end."""
    owner_dtype = pd.CategoricalDtype(categories=order, ordered=True)
    df = df.copy()
    df["Assigned To"] = df["Assigned To"].astype(owner_dtype)
    return df.sort_values("Assigned To").reset_index(drop=True)

#-----------------------------------------------------
# Excel Formatting (openpyxl)
#-----------------------------------------------------

def set_column_width_and_align(ws) -> None:
    """Apply column widths and cell alignment to the Task List sheet."""
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for idx in CENTER_COLUMNS:
            row[idx].alignment = CENTER

#-----------------------------------------------------
# I/O Helpers
#-----------------------------------------------------

def get_output_filename() -> str:
    """Prompt until a valid filename is entered."""
    while True:
        name = input("Enter new task list filename: ").strip()
        if name.endswith(".xlsx"):
            return name
        print("Filename must end with .xlsx")

def write_workbook(df: pd.DataFrame, output_path: Path) -> None:
    """Write Task List workbook to *output_path*."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Task List - curated columns with no index
        df[KEEP_COLUMNS].to_excel(writer, sheet_name="Task List", index=False)

        wb: openpyxl.Workbook = writer.book
        format_task_sheet(wb["Task List"])

    logger.info(f"Saved workbook to {output_path}")

def main() -> None:
    path = WEEKLY_EXPORTS
    if not path.exists():
        raise NotADirectoryError(f"Invalid directory path {path}")

    task_dict = load_task_dict(TASK_DICT_PATH)

    df = load_csvs(path)
    df = assign_owners(df, task_dict)
    df = sort_by_owner(df, CUSTOM_ORDER)

    output_filename = get_output_filename()
    write_workbook(df, path / output_filename)


if __name__ == "__main__":
    main()