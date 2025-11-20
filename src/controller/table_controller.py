from __future__ import annotations

from view.table_view import TableView
from model.table_model import TableModel
from enums.column_enum import ColNum

#모든 기능 로직 처리 

class TableController:
    DEFAULT_PLACEHOLDER_STEPS = 10
    
    def __init__(self):
        self.view: TableView | None = None
        self.model = TableModel()
        
    def set_view(self, view: TableView) -> None:
        self.view = view
        self.view.initialize_placeholder_rows(self.DEFAULT_PLACEHOLDER_STEPS)
        #self._load_placeholder_rows() #Test Rows
        
    def set_seed_data(self):
        seed = self.model.set_seed()        
        self.view.set_load_data(0, seed)
    
    # UI Envents
    def on_cell_clicked(self, row: int, col: int) -> None:
        print(f"Clicked cell {row},{col}")

    def on_cell_changed(self, row: int, col: int) -> None:
        item = self.view.tableWidget.item(row, col)
        if item:
            print(f"Cell changed ({row},{col}) -> {item.text()}")
            
    def handle_close(self):
        print("[Controller] View closed. Saving procedure...")
        self.model.save_procedure()
        
    def change_type(self, row: int, col: int, item: any):
        if not self.view:
            return
        step_index = self._step_index_from_view_row(row)
        step_entity = self.model.update_type(step_index, item)
        if not step_entity:
            return

        displayed_value = item.name if hasattr(item, "name") else item
        self.view.set_combo_value(row, col, displayed_value)
        self.view.mark_step_initialized(row, {"type_": step_entity.type_, "step": step_entity.step})
        
        if step_entity.type_ and step_entity.type_.is_terminal():
            self.model.trim_after(step_index)
            self.view.remove_placeholder_steps_after(row)
        else:
            self.view.ensure_next_step(row)
    
    def change_mode(self, row: int, col: int, item: any):
        pass        
  
    def change_endtype(self, row: int, col: int, item: any):
        pass        
    
    def change_operator(self, row: int, col: int, item: any):
        pass        
    
    def change_report(self, row: int, col: int, item: any):
        pass        
      
    def generate_step(self, row: int):
        self.model.generate_Entitiy(row)

    # Menu Actions
    def copy_step_action(self, row):
        self.model.copy_step(row)

    def paste_step_action(self, row):
        self.model.paste_step(row)
        
    def insert_step_action(self, row):
        step_index = self._step_index_from_view_row(row)
        step_number = self.model.insert_step(step_index)
        if self.view:
            self.view.set_cell_data(row, ColNum.STEP.value , str(step_number))

    def delete_step_action(self, row):
        step_index = self._step_index_from_view_row(row)
        self.model.delete_step(step_index)
        
    # View Test Set
    def _load_placeholder_rows(self, rows: int | None = None):
        total_rows = rows if rows is not None else self.view.tableWidget.rowCount()
        self.view.tableWidget.blockSignals(True)
        for idx in range(total_rows):
            values = [
                str(idx + 1),
                "Type",
                "Mode",
                "0",
                "End Type",
                "=",
                "0",
                "Goto",
                "Report Type",
                "0",
                "Note",
                str(idx + 1),
            ]
            self.view.display_row(idx, values)
        self.view.tableWidget.blockSignals(False)

    def _step_index_from_view_row(self, row: int) -> int:
        if not self.view:
            return row
        return self.view.get_step_index(row)
