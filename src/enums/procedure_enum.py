from enum import Enum, auto
class StepType(Enum):
    Sub = 0
    Charge = auto()
    Discharge = auto()
    Rest = auto()
    Pause = auto()
    Cycle_start = auto()
    Cycle_end = auto()
    EIS = auto()
    DCIR = auto()
    End = auto()

    def is_terminal(self) -> bool:
        return self is StepType.End

class Mode(Enum):
    Current = 0
    Voltage = auto()
    Power = auto()
    Resistance = auto()
    
    def type_to_enum():
        pass

class EndType(Enum):
    Step_time = 0
    Current = auto()
    Voltage = auto()
    Power = auto()
    Amp_Hour = auto()
    Watt_Hour = auto()
    Cycle_Count = auto()
    
    def type_to_enum():
        pass

class Report(Enum):
    Step_time = 0
    Current = auto()
    Voltage = auto()
    Amp_Hour = auto()
    Watt_Hour = auto()
    Dv_dt = auto()
    Di_dt = auto()
    
    def type_to_enum():
        pass

class Operator(Enum):
    Equal = "="
    Less_than = ">="
    More_than = "<="
    Derivative1 = ">d1"
    Derivative2 = "<d1"

    def type_to_enum():
        pass

    @classmethod
    def list(cls):
        """UI용 문자열 리스트 반환"""
        return [op.value for op in cls]

    def __str__(self):
        """출력 시 문자열 기호 그대로 표시"""
        return self.value
