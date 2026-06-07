"""
Create PM Task List from csv files in a provided path
"""
import logging
from tasklist_app.constants import SECTION_CONFIG_PATH, TASK_DICT_PATH, WEEKLY_EXPORTS,CUSTOM_ORDER
from tasklist_app.load import load_csvs, load_section_config, load_task_dict
from tasklist_app.transform import assign_owners, sort_by_owner
from tasklist_app.validate import (
    validate_section_config,
    warn_for_duplicate_task_ids,
    warn_for_unwritten_tasks
)
from tasklist_app.excel_writer import write_workbook
#-----------------------------------------------------
# Logger Setup
#-----------------------------------------------------

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

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

def main() -> None:
    path = WEEKLY_EXPORTS
    if not path.exists():
        raise NotADirectoryError(f"Invalid directory path {path}")

    task_dict = load_task_dict(TASK_DICT_PATH)
    section_config = load_section_config(SECTION_CONFIG_PATH)
    validate_section_config(section_config)

    df = load_csvs(path)
    df = assign_owners(df, task_dict)
    df = sort_by_owner(df, CUSTOM_ORDER)

    warn_for_duplicate_task_ids(df)
    warn_for_unwritten_tasks(df, section_config)

    output_filename = get_output_filename()
    write_workbook(df, section_config,path / output_filename)


if __name__ == "__main__":
    main()
