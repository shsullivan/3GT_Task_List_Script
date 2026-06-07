#-----------------------------------------------------
# Data Loading
#-----------------------------------------------------

import csv
import json
import logging
import pandas as pd
from pathlib import Path
from tasklist_app.models import TaskOwner

logger = logging.getLogger(__name__)

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
