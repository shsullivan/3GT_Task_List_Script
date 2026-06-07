#-----------------------------------------------------
# Data Transformation
#-----------------------------------------------------

import logging
import pandas as pd

logger = logging.getLogger(__name__)


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