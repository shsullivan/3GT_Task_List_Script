#-----------------------------------------------------
# Data/Config Validation
#-----------------------------------------------------

import logging
import pandas as pd
from tasklist_app.models import TaskOwner

logger = logging.getLogger(__name__)

def get_configured_owners(section_config: list[dict]) -> set[str]:
    """Return every owner included in section_config.json"""
    configured_owners = set()

    for section in section_config:
        owner_config = section["owner"]
        owners = owner_config if isinstance(owner_config, list) else [owner_config]
        configured_owners.update(owners)

    return configured_owners

def validate_section_config(section_config: list[dict]) -> None:
    """Validate section_config.json owner values"""
    valid_owners = {member.value for member in TaskOwner}
    configured_owners = get_configured_owners(section_config)

    invalid_owners = configured_owners - valid_owners
    if invalid_owners:
        raise ValueError(f"Unknown owners in section_config.json: {sorted(invalid_owners)}")

def warn_for_duplicate_task_ids(df: pd.DataFrame) -> None:
    """Warn when duplicate task IDs are present in the export data"""
    duplicate_task_ids = sorted(df.loc[df["Task ID"].duplicated(), "Task ID"].dropna().unique())

    if duplicate_task_ids:
        logger.warning("Duplicate task IDs found in export data: %s", duplicate_task_ids)

def warn_for_unwritten_tasks(df: pd.DataFrame, section_config: list[dict]) -> None:
    """Warn for unwritten tasks"""
    configured_owners = get_configured_owners(section_config)
    task_owners = set(df["Assigned To"].dropna().unique())

    missing_owners = task_owners - configured_owners

    if missing_owners:
        missing_tasks = df[df["Assigned To"].isin(missing_owners)]

        logger.warning(
            "%d tasks will not be written because their owners are missing from the section_config.json: %s",
            len(missing_tasks),
            sorted(missing_owners))