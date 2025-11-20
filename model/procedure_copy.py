import enum
from enum import unique

@unique
class Type(enum.Enum):
    End = 0
    Charge = enum.auto()
    Discharge = enum.auto()
    Rest = enum.auto()
    Pause = enum.auto()
    Cycle_start = enum.auto()
    Cycle_end = enum.auto()

    @staticmethod
    def str_to_enum(type_):
        for _type in Type:
            if type_ == _type.name:
                return _type

class Mode(enum.Enum):
    Current = enum.auto()
    Voltage = enum.auto()
    Power = enum.auto()
    Resistance = enum.auto()

    @staticmethod
    def str_to_enum(mode):
        for _mode in Mode:
            if mode == _mode.name:
                return _mode

class Condition(enum.Enum):
    Step_time = enum.auto()
    Current = enum.auto()
    Voltage = enum.auto()
    Power = enum.auto()
    Amp_Hour = enum.auto()
    Watt_Hour = enum.auto()
    Cycle_Count = enum.auto()
    
    @staticmethod
    def str_to_enum(condition):
        for _condition in Condition:
            if condition == _condition.name:
                return _condition

class Operator(enum.Enum):
    Equal = "="
    Less_than = ">="
    More_than = "<="
    Derivative1 = ">d1"
    Derivative2 = "<d1"

class Report(enum.Enum):
    Step_Time = enum.auto()
    Current = enum.auto()
    Voltage = enum.auto()
    Amp_Hour = enum.auto()
    Watt_Hour = enum.auto()
    Dv_dt = enum.auto()
    Di_dt = enum.auto()
    
    @staticmethod
    def str_to_enum(report):
        for _report in Report:
            if report == _report.name:
                return _report

class Step():
    type_: Type | None
    mode: Mode | None
    mode_value: float
    condition: Condition | None
    condition_value: float
    # need 1 or 2 value
    report: list | None
    note: str | None

    def __init__(self) -> None:
        self.mode = 0
        self.condition = 0
        self.report = 0
        self.operator = 0

    def get_types(self) -> list(Type):
        return Type._member_names_

    def get_modes(self) -> list(Mode):
        __get_name = lambda x: x.name
        if self.type_ is Type.Charge:
            return map(__get_name, [Mode(n) for n in [1,2,3]])
        elif self.type_ is Type.Discharge:
            return map(__get_name, [Mode(n) for n in [1,2,3,4]])
        else:
            return None

    def get_conditions(self) -> list(Condition):
        __get_name = lambda x: x.name
        if self.type_ == Type.Charge:
            return map(__get_name, [Condition(n) for n in range(1,6)])
        elif self.type_ == Type.Discharge:
            return map(__get_name, [Condition.Step_time,
                    Condition.Current,
                    Condition.Voltage,
                    Condition.Power,
                    Condition.Ahr,
                    Condition.Whr])
        elif self.type_ == Type.Rest:
            return map(__get_name, [Condition.Step_time,
                    Condition.Current,
                    Condition.Voltage,
                    Condition.Power,
                    Condition.Ahr,
                    Condition.Whr])

    def get_operator(self, index) -> list(Operator):
        __get_value = lambda x: x.value
        if self.condition[index][0] == Condition.Step_time:
            return ["="]
        elif self.condition[index][0] == Condition.Amp_Hour or Condition.Watt_Hour:
            return ["=",">=", "<="]
        else:
            return ["=", ">=", "<=", ">d1", "<d1"]

    def get_reports(self) -> list(Report):
        return Report._member_names_
