import enum
from enum import unique

@unique
class Type(enum.Enum):
    Sub = 0
    Charge = enum.auto()
    Discharge = enum.auto()
    Rest = enum.auto()
    Pause = enum.auto()
    Cycle_start = enum.auto()
    Cycle_end = enum.auto()
    EIS = enum.auto()
    DCIR = enum.auto()
    End = enum.auto()
    
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
    Step_time = enum.auto()
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
    mode: list
    condition: list | None
    # need 1 or 2 value
    report: list | None
    note: str | None

    def __init__(self) -> None:
        self.mode = None
        self.mode_val = None
        self.condition = None
        self.operator = None
        self.condition_val =None
        self.goto = None
        self.report = None
        self.report_val =None
        self.notetext = None
        self.depend_row = None
        
    def end_ch(self):
        return self.type_
    
    def get_types(self) -> list(Type):
        return Type._member_names_

    def get_modes(self) -> list(Mode):
        __get_name = lambda x: x.name
        if self.type_ is Type.Charge:
            return map(__get_name, [Mode(n) for n in [1,2,3]])
        elif self.type_ is Type.Discharge:
            return map(__get_name, [Mode(n) for n in [1,2,3,4]])
        elif self.type_ is Type.DCIR:
            return map(__get_name, [Mode(n) for n in [1]])
        else:
            return None

    def get_conditions(self, type_condition) -> list(Condition):
        __get_name = lambda x: x.name
        if type_condition == None:
            type_condition = self.type_
        if type_condition == Type.Charge:
            return map(__get_name, [Condition(n) for n in range(1,6)])
        elif type_condition == Type.Discharge:
            return map(__get_name, [Condition.Step_time,
                    Condition.Current,
                    Condition.Voltage,
                    Condition.Power,
                    Condition.Amp_Hour,
                    Condition.Watt_Hour])
        elif type_condition == Type.Rest:
            return map(__get_name, [Condition.Step_time,
                    Condition.Current,
                    Condition.Voltage,
                    Condition.Power,
                    Condition.Amp_Hour,
                    Condition.Watt_Hour])
        elif type_condition == Type.Cycle_end:
            return map(__get_name, [Condition.Cycle_Count])
        
        elif type_condition == Type.DCIR:
            return map(__get_name, [Condition.Step_time])
                
        elif type_condition == Type.EIS:
            return print("EIS")
            
    def get_operator(self, index, type_condition) -> list(Operator):
        __get_value = lambda x: x.value
        if self.condition == Condition.Step_time:
            return ["="]
        elif self.condition == Condition.Amp_Hour or self.condition == Condition.Watt_Hour:
            return [ ">=", "<="]
        else:
            return ["=", ">=", "<=", ">d1", "<d1"]

    def get_reports(self) -> list(Report):
        __get_name = lambda x: x.name
        print("report type is : ", self.type_)
        if self.type_ is Type.DCIR:
            return map(__get_name, [Report.Step_time])
        else:
            return Report._member_names_
