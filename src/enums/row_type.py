from enum import Enum, auto

class RowRole(Enum):
    EMPTY = auto()
    STEP = auto()
    SUBSTEP = auto()
    DELIMITER = auto()