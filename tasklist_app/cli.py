#-----------------------------------------------------
# Input Validations for CLI
#-----------------------------------------------------

from datetime import datetime
from pathlib import Path

def prompt_required(prompt: str) -> str:
    """Prompt until the user enters a non-empty value"""
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("This value is required")

def prompt_pm_date() -> str:
    """Prompt for a date and return it formatted for the workbook"""
    while True:
        value = input("Enter PM date (YYYY/MM/DD): ").strip()

        try:
            return datetime.strptime(value, "%Y/%m/%d").strftime("%B %d, %Y")
        except ValueError:
            print("Please enter the date as YYYY/MM/DD. Example: 2026/06/21")

def prompt_output_filename() -> str:
    """Prompt until a valid Excel filename is entered"""
    while True:
        name = input("Enter output filename: ").strip()

        if name.endswith(".xlsx"):
            return name
        else:
            name = name + ".xlsx"
            return name

def collect_run_inputs() -> dict[str, str]:
    """Collect user-provided values for this task list run"""
    line_number = prompt_required("Enter TAM line being PM'd: ")
    pm_week = prompt_required("Enter PM Cycle (16, 32, or 48): ")
    display_date = prompt_pm_date()
    output_filename = prompt_output_filename()

    return {
        "title": f"TAM - {line_number} - {pm_week} Week PM",
        "display_date": display_date,
        "output_filename": output_filename,
    }