#-----------------------------------------------------
# Enums and Custom Types
#-----------------------------------------------------

from enum import Enum


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