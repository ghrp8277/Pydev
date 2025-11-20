
from typing import List, Sequence
import pandas as pd

from model.excel_adapter import ExcelAdapter
from Entities.procedureEntity import ProcedureEntity
from enums.procedure_enum import StepType

#Data Entity Control( Save, Load, Row, Cell )

class TableModel:
    def __init__(self):
        self.excel = ExcelAdapter()
        self.steps: list[ProcedureEntity] = []
        
    #Entity Get/Set
    def get_step_dict(self, step_number: int) -> dict:
        for entity in self.steps:
            if entity.step == step_number:
                return entity.to_dict()
        return {}

    def get_steps_dict(self) -> list[dict]:
        return [entity.to_dict() for entity in self.steps]

    def get_step_list(self, step_number: int) -> list:
        for entity in self.steps:
            if entity.step == step_number:
                return entity.to_list()
        return []

    def get_steps_list(self) -> list[list]:
        return [entity.to_list() for entity in self.steps]
    
    def set_seed(self) -> dict:
        seed_step = ProcedureEntity.seed()
        self.steps.append(seed_step)
        
        return seed_step.to_dict()

    def generate_Entitiy(self, row: int) -> ProcedureEntity:
        new_step = ProcedureEntity(step=row)
        self.steps.append(new_step)
        return new_step               
    
    def _ensure_step_entry(self, step_index: int) -> ProcedureEntity | None:
        if step_index < 0:
            return None

        while step_index >= len(self.steps):
            self.steps.append(ProcedureEntity(step=len(self.steps) + 1))
        return self.steps[step_index]

    def update_type(self, step_index: int, type_value) -> ProcedureEntity | None:
        step = self._ensure_step_entry(step_index)
        if step is None:
            return None

        if not type_value:
            step.type_ = None
            return step

        if isinstance(type_value, StepType):
            step.type_ = type_value
            return step

        value_str = str(type_value)
        try:
            step.type_ = StepType[value_str]
            return step
        except KeyError:
            pass

        try:
            numeric = int(value_str)
            step.type_ = StepType(numeric)
            return step
        except (ValueError, KeyError):
            step.type_ = None
            return step

    def trim_after(self, step_index: int) -> None:
        if step_index < 0:
            self.steps.clear()
            return

        keep_count = min(step_index + 1, len(self.steps))
        del self.steps[keep_count:]
    
    #Step index Control
    def insert_step(self, start_row: int) -> int:
        for index in range(start_row, len(self.steps)):
            if self.steps[index].step is not None:
                self.steps[index].step += 1

        new_step = ProcedureEntity(step=start_row + 1)
        self.steps.insert(start_row, new_step)
        return new_step.step
        
    def delete_step(self, row: int) -> None:
        del self.steps[row]

        for index in range(row, len(self.steps)):
            if self.steps[index].step is not None:
                self.steps[index].step -= 1
            
    def save_procedure(self, file_path: str, steps: List[ProcedureEntity], eiss: Sequence = (), ) -> None:
        step_columns = [
            "Step",
            "Type",
            "Mode",
            "Mode value",
            "End Type",
            "Operator",
            "End Value",
            "Go to",
            "Report Type",
            "Report value",
            "Note",
            "Row",
        ]
        step_rows = [step.to_row() + [None] for idx, step in enumerate(steps)]
        stepdf = pd.DataFrame(step_rows, columns=step_columns)

        if not stepdf.empty:
            stepdf.loc[0, "Row"] = len(stepdf) + 1
        else:
            stepdf.loc[0, "Row"] = 0

        eis_columns = [
            "Mode",
            "Start_Frequency",
            "Stop_Frequency",
            "Amplitude",
            "PointNumber",
            "Row",
        ]
        eis_rows = [
            getattr(e, "to_row", lambda: [])() for e in eiss
        ]
        eisdf = pd.DataFrame(eis_rows, columns=eis_columns)

        eisdf["Row"] = range(1, len(eisdf) + 1)
        if not eisdf.empty:
            eisdf.loc[0, "Row"] = len(eisdf) + 1

        blank_col = pd.DataFrame({"": [""] * max(len(stepdf), len(eisdf))})

        stepdf = stepdf.reindex(range(max(len(stepdf), len(eisdf))))
        eisdf = eisdf.reindex(range(max(len(stepdf), len(eisdf))))

        merged_df = pd.concat([stepdf, blank_col, eisdf], axis=1)

        self.excel.write(file_path, {"Procedure": merged_df})
