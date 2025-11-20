from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from PySide6.QtCore import Qt, Slot, QPoint
from PySide6.QtGui import QCloseEvent, QColor
from PySide6.QtWidgets import (QAbstractItemView, QHeaderView, QTableWidgetItem, QWidget, QMenu, QLineEdit, QComboBox, )

from view.ui import tableWidget_ui
from enums.column_enum import ColNum
from enums.procedure_enum import StepType, Mode, EndType, Operator, Report


from enums.row_type import RowRole
from dataclasses import dataclass
from Entities.procedureEntity import ProcedureEntity

if TYPE_CHECKING:
    from controller.table_controller import TableController
    
COLUMN_WIDGET_MAP = {
    ColNum.TYPE.value: {
        "widget": "combo",
        "items": StepType,
        "func": "change_type"
    },
    ColNum.MODE.value: {
        "widget": "combo",
        "items": Mode,
        "func": "change_mode"
    },
    ColNum.END_TYPE.value: {
        "widget": "combo",
        "items": EndType,
        "func": "change_endtype"
    },
    ColNum.OP.value: {
        "widget": "combo",
        "items": Operator,
        "func": "change_operator"
    },
    ColNum.REPORT_TYPE.value: {
        "widget": "combo",
        "items": Report,
        "func": "change_report"
    }
}

COLUMN_KEY_MAP = {
    "step": ColNum.STEP.value,
    "type_": ColNum.TYPE.value,
    "mode": ColNum.MODE.value,
    "mode_val": ColNum.MODE_VALUE .value,
    "end_type": ColNum.END_TYPE.value,
    "operator": ColNum.OP.value,
    "end_type_val": ColNum.END_TYPE_VAL.value,
    "goto": ColNum.GOTO.value,
    "report": ColNum.REPORT_TYPE.value,
    "report_val": ColNum.REPORT_VALUE .value,
    "note": ColNum.STEP_NOTE .value,
}

COMBO_KEYS = {"type_", "mode", "end_type", "operator", "report"}

#Envet To Controller
#Set view Data

@dataclass
class RowInfo:
    role: RowRole
    data: ProcedureEntity | dict | None = None
    
class TableView(QWidget, tableWidget_ui.TableWidgetUI):
    def __init__(self):
        super().__init__()
        self._controller: Optional["TableController"] = None
        self.setupUi(self)                
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setRowCount(0)
        self._placeholders_initialized = False
        self._active_placeholder_row: Optional[int] = None
        
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.__context_menu)
        self.tableWidget.cellClicked.connect(self._handle_cell_clicked)
        self.tableWidget.cellChanged.connect(self._handle_cell_changed)
        self.__row, self.__col = 0, 0
        self.rows_type: dict[int, RowInfo] = {}

    def _generate_row(self, row: int, row_type: RowRole, data: ProcedureEntity | dict | None = None ):
        self._shift_row_info(row, 1)
        self.rows_type[row] = RowInfo(row_type, data)
        self.tableWidget.insertRow(row)
        if row_type == RowRole.STEP and not self.tableWidget.item(row, ColNum.STEP.value):
            self.tableWidget.setItem(row, ColNum.STEP.value, QTableWidgetItem(""))
        return row

    def _insert_step_block(self, row: int, data: ProcedureEntity | dict | None = None, *, notify_controller: bool = True) -> int:
        step_row = self._generate_row(row, RowRole.STEP, data)
        self._generate_row(step_row + 1, RowRole.DELIMITER, None)
        if notify_controller:
            self._notify_step_created(step_row)
        return step_row

    def _append_step_block(self, data: ProcedureEntity | dict | None = None, *, notify_controller: bool = True) -> int:
        return self._insert_step_block(self.tableWidget.rowCount(), data, notify_controller=notify_controller)

    def _shift_row_info(self, start_row: int, delta: int) -> None:
        if delta == 0 or not self.rows_type:
            return

        updated: dict[int, RowInfo] = {}
        for index in sorted(self.rows_type.keys()):
            info = self.rows_type[index]
            if index >= start_row:
                updated[index + delta] = info
            else:
                updated[index] = info
        self.rows_type = updated
    
    def _notify_step_created(self, row: int) -> None:
        if self._controller:
            self._controller.insert_step_action(row)

    def get_step_index(self, row: int) -> int:
        """Return the zero-based step index corresponding to a view row."""
        count = 0
        for index in sorted(self.rows_type.keys()):
            info = self.rows_type[index]
            if info.role != RowRole.STEP:
                continue
            if index == row:
                return count
            count += 1
        raise ValueError(f"Row {row} is not registered as a step row")
    
    def initialize_placeholder_rows(self, count: int) -> None:
        if self._placeholders_initialized or count <= 0:
            return

        self.tableWidget.blockSignals(True)
        try:
            for _ in range(count):
                self._append_step_block(notify_controller=False)
        finally:
            self.tableWidget.blockSignals(False)
        self._placeholders_initialized = True
        self._activate_next_placeholder()
    
    def _is_step_row(self, row: int) -> bool:
        info = self.rows_type.get(row)
        return bool(info and info.role == RowRole.STEP and info.data is not None)

    def _is_last_step_row(self, row: int) -> bool:
        if not self._is_step_row(row):
            return False

        for index in range(row + 1, self.tableWidget.rowCount()):
            info = self.rows_type.get(index)
            if info and info.role == RowRole.STEP:
                return False
        return True

    def _ensure_trailing_step_exists(self, row: int) -> None:
        if self._is_terminal_step(row):
            self.remove_placeholder_steps_after(row)
            return

        if self._is_last_step_row(row):
            self._append_step_block(notify_controller=bool(self._controller))
        self._activate_next_placeholder()
    
    def ensure_next_step(self, row: int) -> None:
        """외부(Controller)에서 현재 Step이 마지막일 때 다음 Step 준비를 요청할 수 있게 함."""
        self._ensure_trailing_step_exists(row)

    def _is_terminal_step(self, row: int) -> bool:
        info = self.rows_type.get(row)
        if not info or info.role != RowRole.STEP or not info.data:
            return False

        type_value = info.data.get("type_")
        if isinstance(type_value, StepType):
            step_type = type_value
        elif isinstance(type_value, str):
            try:
                step_type = StepType[type_value]
            except KeyError:
                return False
        else:
            return False
        return step_type.is_terminal()

    def mark_step_initialized(self, row: int, data: dict | None = None) -> None:
        info = self.rows_type.get(row)
        if not info or info.role != RowRole.STEP:
            return
        if data is None:
            data = {}
        info.data = data

        step_number = data.get("step")
        if step_number is not None:
            self.set_cell_data(row, ColNum.STEP.value, str(step_number))

    def _find_next_placeholder_row(self) -> Optional[int]:
        for index in sorted(self.rows_type.keys()):
            info = self.rows_type[index]
            if info.role == RowRole.STEP and not info.data:
                return index
        return None

    def _activate_next_placeholder(self) -> None:
        placeholder_row = self._find_next_placeholder_row()
        if placeholder_row == self._active_placeholder_row:
            return

        if self._active_placeholder_row is not None:
            prev_info = self.rows_type.get(self._active_placeholder_row)
            if prev_info and prev_info.data is None:
                self.tableWidget.removeCellWidget(self._active_placeholder_row, ColNum.TYPE.value)
                self.tableWidget.takeItem(self._active_placeholder_row, ColNum.STEP.value)

        if placeholder_row is None:
            self._active_placeholder_row = None
            return

        self.set_cell_attribute(placeholder_row, ColNum.TYPE.value)
        self._active_placeholder_row = placeholder_row

    def remove_placeholder_steps_after(self, row: int) -> None:
        idx = row + 1
        while idx < self.tableWidget.rowCount():
            info = self.rows_type.get(idx)
            if info and info.role == RowRole.STEP and not info.data:
                self._remove_step_block(idx)
                continue
            idx += 1
        self._active_placeholder_row = None

    def _remove_step_block(self, step_row: int) -> None:
        # remove delimiter first to avoid index shift messing with step row removal
        delimiter_row = step_row + 1
        if delimiter_row < self.tableWidget.rowCount():
            self._remove_row(delimiter_row)
        self._remove_row(step_row)

    def _remove_row(self, row: int) -> None:
        if row in self.rows_type:
            self.rows_type.pop(row)
        self.tableWidget.removeRow(row)
        self._shift_row_info(row + 1, -1)
        
    def set_controller(self, controller: "TableController") -> None:
        self._controller = controller
        self._activate_next_placeholder()

    def display_row(self, row_index: int, values: list[str]) -> None:
        for col, text in enumerate(values):
            item = self.tableWidget.item(row_index, col)
            if not item:
                item = QTableWidgetItem()
                self.tableWidget.setItem(row_index, col, item)
            item.setText(text)
            
    def set_load_data(self, row: int, data: dict):
        print(data)
        data.pop('depend_row', None)
        if row not in self.rows_type:
            row = self._insert_step_block(row, notify_controller=False)

        self.tableWidget.blockSignals(True)
        try:
            for key, value in data.items():
                col = COLUMN_KEY_MAP[key]

                if key == "step":
                    self.set_cell_data(row, col, "" if value is None else str(value))
                    continue

                if key in COMBO_KEYS:
                    items = self._build_combo_items(COLUMN_WIDGET_MAP[col]["items"])
                    func_name = COLUMN_WIDGET_MAP[col]["func"]
                    change_func = getattr(self._controller, func_name, None) if self._controller else None

                    widget = self._generate_combo_box(row, col, items, change_func)
                    if value is not None:
                        widget.setCurrentText(str(value))
                    self.tableWidget.setCellWidget(row, col, widget)
                else:
                    widget = self._generate_line_edit()
                    widget.setText("" if value is None else str(value))
                    self.tableWidget.setCellWidget(row, col, widget)
        finally:
            self.tableWidget.blockSignals(False)

        row_info = self.rows_type.get(row)
        if row_info:
            row_info.data = data
        self._ensure_trailing_step_exists(row)

    def set_cell_data(self, row: int, col: int, value: str=" "):
        item = QTableWidgetItem(str(value))
        #item.setBackground(QColor(25, 25, 25))
        self.tableWidget.setItem(row, col, item)
        
    def set_combo_value(self, row: int, col: int, value: str | None):
        widget = self.tableWidget.cellWidget(row, col)
        if isinstance(widget, QComboBox):
            block_state = widget.blockSignals(True)
            widget.setCurrentText("" if value is None else str(value))
            widget.blockSignals(block_state)
    def set_cell_attribute(self, row: int, col: int):
        conf = COLUMN_WIDGET_MAP.get(col)

        if conf and conf["widget"] == "combo":
            items = self._build_combo_items(conf["items"])
            func_name = conf["func"]
            change_func = getattr(self._controller, func_name, None) if self._controller else None

            widget = self._generate_combo_box(row, col, items, change_func)
            self.tableWidget.setCellWidget(row, col, widget)
        else:
            widget = self._generate_line_edit()
            self.tableWidget.setCellWidget(row, col, widget)

    def _build_combo_items(self, enum_cls) -> list[str]:
        items = [name for name in enum_cls.__members__.keys()]
        if enum_cls is StepType and "Sub" in items:
            items = [name for name in items if name != "Sub"]
        return items

    def _generate_combo_box(self, row: int, col: int, items: list, change_func=None) -> QComboBox:
        combo = QComboBox()
        combo.addItems(items)
        block_state = combo.blockSignals(True)
        combo.setCurrentIndex(-1)
        combo.blockSignals(block_state)
        if change_func:
            combo.currentTextChanged.connect(
                lambda value, f=change_func: f(row, col, value)
            )
        return combo
    
    def _generate_line_edit(self) -> QLineEdit:
        line = QLineEdit()
        line.textChanged.connect(self._on_line_changed)
        return line
    
    def _on_combo_changed(self, text):
        print("Combo updated:", text)

    def _on_line_changed(self, text):
        print("Line updated:", text)

    def _handle_cell_clicked(self, row: int, col: int) -> None:
        self.__row, self.__col = row, col
        if self._controller:
            self._controller.on_cell_clicked(row, col)

    def _handle_cell_changed(self, row: int, col: int) -> None:
        if self._controller:
            self._controller.on_cell_changed(row, col)
            
    # Context menu
    @Slot(QPoint)
    def __context_menu(self, pos):
        menu = QMenu()
        insert_action = menu.addAction("Insert Step")
        delete_action = menu.addAction("Delete Step")
        copy_action = menu.addAction("Copy Step")
        paste_action = menu.addAction("Paste Step")
        menu.addSeparator()
        
        clear_cell_action = menu.addAction("Clear Cell")
        clear_line_action = menu.addAction("Clear Line")
        
        action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        row = self.tableWidget.indexAt(pos).row()

        if not self._controller:
            return

        if action == copy_action:
            pass
        elif action == paste_action:
            pass
        elif action == insert_action:
            step_row = self._insert_step_block(row)
            self._ensure_trailing_step_exists(step_row)
        elif action == delete_action:
            self._controller.delete_step_action(row)
        elif action == clear_cell_action:
            pass
        elif action == clear_line_action:
            pass
            
    def closeEvent(self, event: QCloseEvent):
        if self._controller:
            self._controller.handle_close()
        super().closeEvent(event)
