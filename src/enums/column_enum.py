from enum import Enum, auto 

class ColNum(Enum):
    STEP = 0
    TYPE = auto()
    MODE = auto()
    MODE_VALUE = auto()
    END_TYPE = auto()
    OP = auto()
    END_TYPE_VAL = auto()
    GOTO = auto()
    REPORT_TYPE = auto()
    REPORT_VALUE = auto()
    STEP_NOTE = auto()