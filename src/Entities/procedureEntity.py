
from dataclasses import dataclass
from typing import Optional, Any

from enums.procedure_enum import StepType, Mode, EndType, Report, Operator

def _to_enum(enum_class, raw_value: Any):
    if raw_value is None or raw_value == "":
        return None

    if isinstance(raw_value, enum_class):
        return raw_value

    value_as_string = str(raw_value)

    try:
        return enum_class[value_as_string]
    except Exception:
        pass
    try:
        integer_value = int(float(raw_value))
        for member in enum_class:
            if member.value == integer_value:
                return member
    except Exception:
        pass

    return None

@dataclass
class ProcedureEntity:
    step: Optional[int] = None
    type_: Optional[StepType] = None
    mode: Optional[Mode] = None
    mode_val: Optional[float] = None
    end_type: Optional[EndType] = None
    operator: Optional[Operator] = None
    end_type_val: Optional[float] = None
    goto: Optional[int] = None
    report: Optional[Report] = None
    report_val: Optional[float] = None
    note: Optional[str] = None
    depend_row: Optional[int] = None

    def to_dict(self) -> dict:
        return {
            "step": self.step,
            "type_": self.type_.name if self.type_ else None,
            "mode": self.mode.name if self.mode else None,
            "mode_val": self.mode_val,
            "end_type": self.end_type.name if self.end_type else None,
            "operator": self.operator,
            "end_type_val": self.end_type_val,
            "goto": self.goto,
            "report": self.report.name if self.report else None,
            "report_val": self.report_val,
            "note": self.note,
            "depend_row": self.depend_row
        }

    def to_list(self) -> list:
        return [
            self.step,
            self.type_.name if self.type_ else None,
            self.mode.name if self.mode else None,
            self.mode_val,
            self.end_type.name if self.end_type else None,
            self.operator,
            self.end_type_val,
            self.goto,
            self.report.name if self.report else None,
            self.report_val,
            self.note,
            self.depend_row
        ]

    @staticmethod
    def seed():
        return ProcedureEntity(
            step=1, 
            type_=StepType.Rest,
            mode=None,
            mode_val=None,
            end_type=EndType.Step_time,
            operator="=",
            end_type_val=3.0,
            goto=2,
            report=Report.Step_time,
            report_val=1.0,
            note=None,
        )