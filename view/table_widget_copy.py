# pylint: disable = no-name-in-module

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional, List
from abc import ABC, abstractmethod
import enum
import copy
import math

from PySide6.QtGui import QCloseEvent, QAction
from PySide6.QtWidgets import (
    QWidget, QFileDialog, QTableWidgetItem, QMenu, QAbstractItemView,
    QComboBox, QLineEdit
)
from PySide6.QtCore import Qt, Slot, QPoint

from view import ui_BuildTestTableWidget
from model import procedure
from model import file
import pandas as pd


# ==================== 공용 Enum / Item ==================== #

class ZeroBasedEnum(enum.Enum):
    def _generate_next_value_(name, start, count, last_values):
        return count  # 0부터 시작


class ColNum(ZeroBasedEnum):
    STEP = enum.auto()
    TYPE = enum.auto()
    MODE = enum.auto()
    MODE_VALUE = enum.auto()
    END_TYPE = enum.auto()
    OP = enum.auto()
    OP_VALUE = enum.auto()
    GOTO = enum.auto()
    REPORT_TYPE = enum.auto()
    REPORT_VALUE = enum.auto()
    STEP_NOTE = enum.auto()


class TableItem(QTableWidgetItem):
    function = None


class Box_method_item:
    type: None
    mode: None
    endtype: None
    operator: None
    retype: None


# ==================== 메멘토 / 커맨드 패턴 ==================== #

@dataclass
class TableMemento:
    """
    TableWidget의 '모델 상태' 스냅샷.
    - UI 위젯(QTableWidget) 자체는 저장하지 않고,
      step_list / box_list / eis_parameters 같은 핵심 데이터만 저장.
    """
    step_list: list
    box_list: list
    eis_parameters: list


class Command(ABC):
    """
    Undo/Redo를 지원하는 커맨드 기본 클래스.
    - execute()  : 명령 실행
    - undo()     : 이전 상태로 롤백
    - redo()     : 다시 실행 (필요 시)
    """

    def __init__(self, table: "TableWidget") -> None:
        self.table = table
        self._before: Optional[TableMemento] = None
        self._after: Optional[TableMemento] = None

    def execute(self) -> None:
        self._before = self.table.create_memento()
        self._do()
        self._after = self.table.create_memento()

    def undo(self) -> None:
        if self._before:
            self.table.restore_memento(self._before)

    def redo(self) -> None:
        if self._after:
            self.table.restore_memento(self._after)

    @abstractmethod
    def _do(self) -> None:
        """실제 명령 실행 로직"""


class InsertStepCommand(Command):
    def __init__(self, table: "TableWidget", row: int) -> None:
        super().__init__(table)
        self.row = row

    def _do(self) -> None:
        self.table.insert_step_above(self.row)


class DeleteStepBlockCommand(Command):
    def __init__(self, table: "TableWidget", row: int) -> None:
        super().__init__(table)
        self.row = row

    def _do(self) -> None:
        self.table.delete_step_block(self.row)


class PasteStepCommand(Command):
    def __init__(self, table: "TableWidget", row: int) -> None:
        super().__init__(table)
        self.row = row

    def _do(self) -> None:
        if self.table.paste_list:
            # 첫 번째 Step만 메인으로 붙이는 기존 로직 활용
            self.table.paste_step_above(self.row, self.table.paste_list[0][1])


# ==================== 초기화 전략 패턴 ==================== #

class StepTableInitializer(ABC):
    @abstractmethod
    def init(self, table: "TableWidget") -> None:
        """TableWidget의 상태를 초기화한다."""
        ...


class NewProcedureInitializer(StepTableInitializer):
    def init(self, table: "TableWidget") -> None:
        table.new_step(0)
        table.new_step(2)
        table.new_cell(2, ColNum.TYPE.value, method_name=table.generate_combo_box)

        table.reset_step_index()

        init_rest = ["Rest", "Step_time", "=", 3, "2", "Step_time", 1]
        for col, item in enumerate(init_rest):
            if col < ColNum.TYPE.value:
                set_col = col + 1
            else:
                set_col = col + 3

            if set_col in (
                ColNum.TYPE.value,
                ColNum.END_TYPE.value,
                ColNum.OP.value,
                ColNum.REPORT_TYPE.value,
            ):
                table.new_cell(0, set_col, text=item, method_name=table.generate_combo_box)
                table._TableWidget__change_combo_box(
                    0, set_col, str(item), table.find_method_box(set_col)
                )
            else:
                table.new_cell(0, set_col, text=item, method_name=table.generate_value_cell)
                table._TableWidget__change_value(0, set_col, item)


class FileLoadInitializer(StepTableInitializer):
    def __init__(self, file_name: str) -> None:
        self.file_name = file_name

    def init(self, table: "TableWidget") -> None:
        load_step_list, temp_eis = file.load_full_excel(self.file_name)

        if load_step_list:
            del load_step_list[0]
        if temp_eis:
            del temp_eis[0]

        # EIS 경로 매칭
        for eis_item in temp_eis:
            eis_path = table.search_eis_path(eis_item[0], load_step_list)
            eis_item.append(eis_path)
            table.eis_parameters.append(eis_item)

        load_step_package = []
        package_main_index = []

        # 메인 스텝 시작 인덱스 추출
        for idx, row in enumerate(load_step_list):
            if row:
                del row[-1]
            if row and not pd.isna(row[0]):
                package_main_index.append(idx)

        if not package_main_index:
            package_main_index = [0]
        package_main_index.append(len(load_step_list))

        for i, start in enumerate(package_main_index[:-1]):
            end = package_main_index[i + 1]
            load_step_package.append(load_step_list[start:end])

        set_row = 0
        for pkg in load_step_package:
            table.load_step(pkg, set_row)
            set_row += (len(pkg) + 1)


# ==================== TableWidget ==================== #

class TableWidget(QWidget, ui_BuildTestTableWidget.Ui_BuildTestTableWidget):
    def __init__(self, status_widgets=None, file_name=None) -> None:
        super().__init__()
        self.setupUi(self)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.__row = 0
        self.__col = 0
        self.file = file_name
        self.status_widgets = status_widgets
        self.init_rest = ["Rest", "Step_time", "=", 3, "2", "Step_time", 1]
        self.step_list: list[list[Any]] = []
        self.box_list: list[list[Any]] = []
        self.load_step_list = []
        self.eis_parameters: list[list[Any]] = []
        self.sub_step_cnt = 0

        self.paste_list = []
        self.paste_box_list = []

        # === Undo/Redo 스택 ===
        self._undo_stack: List[Command] = []
        self._redo_stack: List[Command] = []

        # 시그널 설정
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.__context_menu)
        self.tableWidget.cellClicked.connect(self.__click_cell)
        self.tableWidget.cellChanged.connect(self.__changed_cell)

        # 초기화 전략 선택
        if self.file is None:
            initializer: StepTableInitializer = NewProcedureInitializer()
        else:
            initializer = FileLoadInitializer(self.file)

        try:
            initializer.init(self)
        except Exception as e:
            print("load error:", e)

    # ==================== 메멘토 관련 ==================== #

    def create_memento(self) -> TableMemento:
        """현재 모델 상태(step_list, box_list, eis_parameters)를 스냅샷으로 저장."""
        return TableMemento(
            step_list=copy.deepcopy(self.step_list),
            box_list=copy.deepcopy(self.box_list),
            eis_parameters=copy.deepcopy(self.eis_parameters),
        )

    def restore_memento(self, m: TableMemento) -> None:
        """
        메멘토로부터 모델 상태 복원 + 테이블 UI 재구성.
        - 완전한 형상관리처럼 '모델 기준으로 UI를 재빌드'하는 방식.
        """
        self.step_list = copy.deepcopy(m.step_list)
        self.box_list = copy.deepcopy(m.box_list)
        self.eis_parameters = copy.deepcopy(m.eis_parameters)

        # UI 클리어
        self.tableWidget.blockSignals(True)
        try:
            self.tableWidget.setRowCount(0)
            # row 인덱스를 기준으로 다시 행 생성
            for row, step in sorted(self.step_list, key=lambda x: x[0]):
                if row >= self.tableWidget.rowCount():
                    self.tableWidget.insertRow(self.tableWidget.rowCount())

                # Step 번호 셀
                if hasattr(step, "type_") and step.type_ == procedure.Type.Sub:
                    self.new_cell(row, ColNum.STEP.value, text=" ")
                else:
                    # reset_step_index에서 다시 정리되므로 임시값
                    self.new_cell(row, ColNum.STEP.value, text=" ")

                # TYPE
                if hasattr(step, "type_") and step.type_ is not None:
                    tname = step.type_.name
                    self.new_cell(row, ColNum.TYPE.value, text=tname, method_name=self.generate_combo_box)
                    self.change_type(row, tname, isLoad=1)

                # MODE
                if hasattr(step, "mode") and step.mode is not None:
                    mname = step.mode.name
                    self.new_cell(row, ColNum.MODE.value, text=mname, method_name=self.generate_combo_box)
                    self.change_mode(row, mname, isLoad=1)

                # END_TYPE
                if hasattr(step, "condition") and step.condition is not None:
                    cname = step.condition.name
                    self.new_cell(row, ColNum.END_TYPE.value, text=cname, method_name=self.generate_combo_box)
                    self.change_condition(row, cname, isLoad=1)

                # OP
                if hasattr(step, "operator") and step.operator is not None:
                    oname = str(step.operator)
                    self.new_cell(row, ColNum.OP.value, text=oname, method_name=self.generate_combo_box)
                    self.change_operator(row, oname, isLoad=1)

                # REPORT_TYPE
                if hasattr(step, "report") and step.report is not None:
                    rname = step.report.name
                    self.new_cell(row, ColNum.REPORT_TYPE.value, text=rname, method_name=self.generate_combo_box)
                    self.change_report(row, rname, isLoad=1)

                # VALUE 계열
                if getattr(step, "mode_val", None) is not None:
                    self.new_cell(row, ColNum.MODE_VALUE.value, text=str(step.mode_val),
                                  method_name=self.generate_value_cell)
                    self.__change_value(row, ColNum.MODE_VALUE.value, step.mode_val)
                if getattr(step, "condition_val", None) is not None:
                    self.new_cell(row, ColNum.OP_VALUE.value, text=str(step.condition_val),
                                  method_name=self.generate_value_cell)
                    self.__change_value(row, ColNum.OP_VALUE.value, step.condition_val)
                if getattr(step, "goto", None) is not None:
                    self.new_cell(row, ColNum.GOTO.value, text=str(step.goto),
                                  method_name=self.generate_value_cell)
                    self.__change_value(row, ColNum.GOTO.value, step.goto)
                if getattr(step, "report_val", None) is not None:
                    self.new_cell(row, ColNum.REPORT_VALUE.value, text=str(step.report_val),
                                  method_name=self.generate_value_cell)
                    self.__change_value(row, ColNum.REPORT_VALUE.value, step.report_val)
                if getattr(step, "notetext", None) is not None:
                    self.new_cell(row, ColNum.STEP_NOTE.value, text=str(step.notetext),
                                  method_name=self.generate_value_cell)
                    self.__change_value(row, ColNum.STEP_NOTE.value, step.notetext)

            # Step 번호 / goto 재계산
            self.reset_step_index()
            self.reset_goto()
        finally:
            self.tableWidget.blockSignals(False)

    # ==================== Command 실행 헬퍼 ==================== #

    def execute_command(self, cmd: Command) -> None:
        cmd.execute()
        self._undo_stack.append(cmd)
        self._redo_stack.clear()

    def undo(self) -> None:
        if not self._undo_stack:
            return
        cmd = self._undo_stack.pop()
        cmd.undo()
        self._redo_stack.append(cmd)

    def redo(self) -> None:
        if not self._redo_stack:
            return
        cmd = self._redo_stack.pop()
        cmd.redo()
        self._undo_stack.append(cmd)

    # ==================== 기존 메서드들 (원본에서 가져온 부분) ==================== #

    def load_step(self, load_step_list, set_row):
        prev_main_row = None

        for row_offset, repeat_step in enumerate(load_step_list):
            insert_row = set_row + row_offset
            self.new_step(insert_row)

            for col, value in enumerate(repeat_step):
                if value is None or (isinstance(value, float) and math.isnan(value)):
                    value = " "

                if col == ColNum.STEP.value:
                    continue

                if col == ColNum.TYPE.value:
                    text_val = str(value).strip()
                    func = self.find_method_box(col)

                    self.new_cell(insert_row, col, text=text_val, method_name=self.generate_combo_box)

                    if text_val in ["", " ", "nan", "None"]:
                        # Sub Step
                        step_obj = self.find_step_obj(insert_row)
                        step_obj.type_ = procedure.Type.Sub
                        step_obj.depend_row = prev_main_row
                        self.__change_combo_box(insert_row, col, "Sub", func, isLoad=1)
                    else:
                        # Main Step
                        prev_main_row = insert_row
                        self.__change_combo_box(insert_row, col, text_val, func, isLoad=1)
                    continue

                elif col in (
                    ColNum.MODE.value,
                    ColNum.END_TYPE.value,
                    ColNum.OP.value,
                    ColNum.REPORT_TYPE.value,
                ):
                    self.new_cell(insert_row, col, text=str(value), method_name=self.generate_combo_box)
                    func = self.find_method_box(col)
                    if str(value).strip() not in ["", " ", "nan"]:
                        self.__change_combo_box(insert_row, col, str(value), func, isLoad=1)

                elif col in (
                    ColNum.MODE_VALUE.value,
                    ColNum.OP_VALUE.value,
                    ColNum.GOTO.value,
                    ColNum.REPORT_VALUE.value,
                    ColNum.STEP_NOTE.value,
                ):
                    self.new_cell(insert_row, col, text=str(value), method_name=self.generate_value_cell)
                    if str(value).strip() not in ["", " ", "nan"]:
                        self.__change_value(insert_row, col, value)

        self.reset_step_index()

        for step in self.step_list:
            if getattr(step[1], "type_", None) == procedure.Type.Sub:
                dep_idx = self.find_step_index(step[1].depend_row)
                if dep_idx is not None:
                    step[1].depend_row = self.step_list[dep_idx][0]
                else:
                    step[1].depend_row = 0

    def search_eis_path(self, step_num: int, step_list: list[list]) -> Optional[str]:
        for step in step_list:
            if not step or len(step) < 3:
                continue
            if str(step[0]) == str(step_num) and str(step[1]).upper() == "EIS":
                return step[2]
        return None

    def find_method_box(self, col):
        select_method = None
        if col == ColNum.TYPE.value:
            select_method = self.change_type
        elif col == ColNum.MODE.value:
            select_method = self.change_mode
        elif col == ColNum.END_TYPE.value:
            select_method = self.change_condition
        elif col == ColNum.OP.value:
            select_method = self.change_operator
        elif col == ColNum.REPORT_TYPE.value:
            select_method = self.change_report
        return select_method

    def __click_cell(self, row, col):
        self.tableWidget.removeCellWidget(self.__row, self.__col)

        if not self.box_list or not self.box_list[0]:
            print("null box_list_item list create")
        else:
            now_box_method = None
            for row_method in self.box_list:
                if row_method[0] == row:
                    now_box_method = row_method[1]

        cell_item_call = self.tableWidget.item(row, col)

        if cell_item_call and cell_item_call.function is not None:
            cell_item_call.function(row, col, self.tableWidget.item(row, col).text(), box_method=now_box_method)

        self.__row = row
        self.__col = col
        print("select cell:", row, col)

    def __changed_cell(self, row, col):
        stepType = None
        if self.__row == row and self.__col == col:
            step = self.find_step_obj(row)
            if hasattr(step, "type_"):
                stepType = step.type_

            if stepType == procedure.Type.EIS:
                find_step_number = self.find_step_number(row)
                print("EIS 설정창 열기 요청")
                parent_window = self.window()
                if hasattr(parent_window, "open_eis_window"):
                    parent_window.open_eis_window(find_step_number)
                self.change_condition(row, " ", 0)

            if col == ColNum.END_TYPE.value:
                if stepType == procedure.Type.Sub:
                    if not self.find_step_obj(row + 1):
                        self.sub_step(row)
                elif stepType == procedure.Type.Charge or stepType == procedure.Type.Discharge:
                    self.sub_step(row, 1)

    def find_step_obj(self, row):
        for i in self.step_list:
            if i[0] == row:
                return i[1]
        return None

    def reset_goto(self):
        # 1) Main 먼저 계산
        for row, step in self.step_list:
            if not hasattr(step, "type_") or step.type_ == procedure.Type.Sub:
                continue

            step_no = self._safe_step_number(row)
            if step_no is None:
                continue

            goto = step_no + 1
            self.__change_value(row, ColNum.GOTO.value, goto)

        # 2) Sub를 부모와 동기화
        for row, step in self.step_list:
            if not hasattr(step, "type_") or step.type_ != procedure.Type.Sub:
                continue

            depend_row = getattr(step, "depend_row", None)
            if depend_row is None:
                continue

            parent_idx = self.find_step_index(depend_row)
            if parent_idx is None:
                continue

            parent_row, _ = self.step_list[parent_idx]
            parent_goto_txt = self._cell_text(parent_row, ColNum.GOTO.value)
            if parent_goto_txt:
                self.__change_value(row, ColNum.GOTO.value, parent_goto_txt)

    def new_step(self, row):
        new_step = [row, procedure.Step()]
        new_box_method = [row, Box_method_item()]

        self.step_list.append(new_step)
        self.box_list.append(new_box_method)

        self.step_list.sort(key=lambda x: x[0])
        self.box_list.sort(key=lambda x: x[0])

    def update_row(self, Current_row):
        for row_index in range(Current_row, len(self.step_list)):
            self.step_list[row_index][0] += 1
            self.box_list[row_index][0] += 1

    def sub_step(self, row, new_sub=None):
        sub_step_row = row + 1
        self.new_step(sub_step_row)

        update_start_index = self.find_step_index(sub_step_row)
        self.update_row(update_start_index + 1)

        step = self.step_list[update_start_index][1]
        step.type_ = procedure.Type.str_to_enum("Sub")

        if new_sub:
            step.depend_row = row
        else:
            step.depend_row = self.find_step_obj(row).depend_row

        for i in range(update_start_index + 1, len(self.step_list)):
            if self.step_list[i][1].__dict__.__contains__("type_"):
                if self.step_list[i][1].type_ == procedure.Type.Sub:
                    self.step_list[i][1].depend_row += 1

        self.tableWidget.insertRow(sub_step_row)
        sub_combo_list = [4, 8]
        for combo_col in sub_combo_list:
            self.new_cell(sub_step_row, combo_col, method_name=self.generate_combo_box)

        self.sub_step_cnt += 1

    def new_cell(self, row, col, text: str = " ", method_name=None):
        if pd.isna(text):
            text = " "
        cell = TableItem(text)
        cell.function = method_name
        self.tableWidget.removeCellWidget(row, col)
        self.tableWidget.setItem(row, col, cell)

    def generate_combo_box(self, row, col, text=" ", box_method=None, isLoad=None):
        step_index = self.find_step_index(row)
        if step_index is None or step_index >= len(self.step_list):
            return

        step = self.step_list[step_index][1]

        if not hasattr(step, "type_"):
            step.type_ = None

        combo = QComboBox()
        combo.setPlaceholderText(text)
        combo.setCurrentIndex(-1)

        type_condition = None
        if getattr(step, "type_", None) == procedure.Type.Sub:
            depend_row = getattr(step, "depend_row", None)
            if depend_row is not None:
                parent_idx = self.find_step_index(depend_row)
                if parent_idx is not None:
                    parent_step = self.step_list[parent_idx][1]
                    if hasattr(parent_step, "type_"):
                        type_condition = parent_step.type_

        col_config = None

        try:
            if col == ColNum.TYPE.value:
                items = [t for t in procedure.Type._member_names_ if t != "Sub"]
                col_config = {"items": items, "func": self.change_type, "attr": "type"}
            elif col == ColNum.MODE.value:
                items = list(map(str, step.get_modes()))
                col_config = {"items": items, "func": self.change_mode, "attr": "mode"}
            elif col == ColNum.END_TYPE.value:
                items = list(map(str, step.get_conditions(type_condition)))
                col_config = {"items": items, "func": self.change_condition, "attr": "endtype"}
            elif col == ColNum.OP.value:
                items = list(map(str, step.get_operator(step_index, type_condition)))
                col_config = {"items": items, "func": self.change_operator, "attr": "operator"}
            elif col == ColNum.REPORT_TYPE.value:
                items = list(map(str, step.get_reports()))
                col_config = {"items": items, "func": self.change_report, "attr": "retype"}
        except Exception:
            col_config = None

        if not col_config or not col_config.get("items"):
            return

        items = col_config["items"]
        change_func = col_config["func"]
        attr_name = col_config["attr"]

        if text in items:
            items.remove(text)

        combo.addItems(items)
        combo.currentTextChanged.connect(
            lambda value: self.__change_combo_box(row, col, value, change_func, isLoad)
        )

        if box_method:
            setattr(box_method, attr_name, combo)

        self.tableWidget.setCellWidget(row, col, combo)

    def __change_combo_box(self, row, col, text, change_func, isLoad=None):
        change_func(row, text, isLoad)
        if col == ColNum.TYPE.value and text == "Sub":
            return
        item = self.tableWidget.item(row, col)
        (item.setText(text) if item else self.new_cell(row, col, text=text, method_name=self.generate_value_cell))

    def change_type(self, row, text, isLoad=None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.type_ = procedure.Type.str_to_enum(text)

        if isLoad:
            return
        self.reset_val(row)

        combo_cols = [ColNum.MODE.value, ColNum.END_TYPE.value, ColNum.REPORT_TYPE.value, ColNum.OP.value]
        value_cols = [ColNum.MODE_VALUE.value, ColNum.OP_VALUE.value, ColNum.REPORT_VALUE.value]

        for col in combo_cols:
            self.new_cell(row, col, method_name=self.generate_combo_box)
        for col in value_cols:
            self.new_cell(row, col, method_name=self.generate_value_cell)

        self.new_cell(row, ColNum.GOTO.value, method_name=self.generate_value_cell)
        self.new_cell(row, ColNum.STEP_NOTE.value, method_name=self.generate_value_cell)

        if step.type_ == procedure.Type.Cycle_start:
            self.reset_goto()

        elif step.type_ == procedure.Type.Cycle_end:
            self.__set_combo_and_apply(row, ColNum.END_TYPE.value, "Cycle_Count", self.change_condition)
            self.__set_combo_and_apply(row, ColNum.OP.value, "=", self.change_operator)
            self.__set_combo_and_apply(row, ColNum.REPORT_TYPE.value, "Step_time", self.change_report)
            self.reset_goto()

        elif step.type_ == procedure.Type.DCIR:
            self.__set_combo_and_apply(row, ColNum.MODE.value, "Current", self.change_mode)
            self.__set_combo_and_apply(row, ColNum.REPORT_TYPE.value, "Step_time", self.change_report)
            self.__set_combo_and_apply(row, ColNum.END_TYPE.value, "Step_time", self.change_condition)
            self.__set_combo_and_apply(row, ColNum.OP.value, "=", self.change_operator)
            self.new_cell(row, ColNum.OP_VALUE.value, text=" ", method_name=self.generate_value_cell)
            self.new_cell(row, ColNum.REPORT_VALUE.value, text=" ", method_name=self.generate_value_cell)
            self.new_cell(row, ColNum.STEP_NOTE.value, text=" ", method_name=self.generate_value_cell)

        if (
            self.step_list[-1][0] == row
            and text != "End"
            and self.step_list[-1][1].__dict__.__contains__("type_")
        ):
            self.new_step(self.step_list[-1][0] + 2)
            self.new_cell(self.step_list[-1][0], ColNum.TYPE.value, method_name=self.generate_combo_box)
            self.reset_step_index()

    def __set_combo_and_apply(self, row, col, value, change_func):
        self.new_cell(row, col, text=value, method_name=self.generate_combo_box)
        self.__change_combo_box(row, col, str(value), self.find_method_box(col))
        change_func(row, value, 0)

    def change_mode(self, row, text, isLoad=None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.mode = procedure.Mode.str_to_enum(text)
        if not isLoad:
            self.new_cell(row, ColNum.MODE_VALUE.value, method_name=self.generate_value_cell)

    def reset_val(self, row):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.mode = None
        step.mode_val = None
        step.condition = None
        step.operator = None
        step.condition_val = None
        step.report = None
        step.report_val = None
        step.notetext = None

    def change_condition(self, row, text, isLoad=None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.condition = procedure.Condition.str_to_enum(text)

        if isLoad:
            return
        else:
            insert_goto = None
            if self.tableWidget.item(row, ColNum.STEP.value):
                step_item = self.tableWidget.item(row, ColNum.STEP.value).text()
                if step_item != " ":
                    insert_goto = int(step_item) + 1
            if step.type_ == procedure.Type.Sub:
                find_root_step = self.step_list[self.find_step_index(step.depend_row)][1]
                step.goto = find_root_step.goto
                insert_goto = step.goto
            else:
                step.goto = str(insert_goto)
            self.new_cell(row, ColNum.OP.value, method_name=self.generate_combo_box)
            self.new_cell(row, ColNum.OP_VALUE.value, method_name=self.generate_value_cell)
            self.new_cell(row, ColNum.GOTO.value, text=str(insert_goto), method_name=self.generate_value_cell)

    def change_operator(self, row, text, isLoad=None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.operator = text
        print("여기 op 는 : ", step_index, row)

    def change_report(self, row, text, isLoad=None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.report = procedure.Report.str_to_enum(text)
        if not isLoad:
            self.new_cell(row, ColNum.REPORT_VALUE.value, method_name=self.generate_value_cell)

    def generate_value_cell(self, row, col, text=" ", box_method=None):
        value_cell = QLineEdit(text)
        value_cell.textChanged.connect(
            lambda _text: self.__change_value(row, col, _text)
        )
        self.tableWidget.setCellWidget(row, col, value_cell)

    def reset_step_index(self):
        seq = 1
        for row, step in sorted(self.step_list, key=lambda x: x[0]):
            if hasattr(step, "type_") and step.type_ == procedure.Type.Sub:
                self._set_step_text(row, "")
            else:
                self._set_step_text(row, seq)
                seq += 1

    def __change_value(self, row, col, text):
        try:
            print(row, col, text)
            if col == ColNum.STEP.value:
                return
            elif col == ColNum.GOTO.value:
                value = text
            elif col == ColNum.STEP_NOTE.value:
                value = str(text)
            else:
                value = float(text)
            self.tableWidget.item(row, col).setText(str(value))
        except Exception as e:
            print(e)
            return

        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]

        if col == ColNum.MODE_VALUE.value:
            step.mode_val = value
        elif col == ColNum.OP_VALUE.value:
            step.condition_val = value
        elif col == ColNum.GOTO.value:
            step.goto = value
        elif col == ColNum.REPORT_VALUE.value:
            step.report_val = value
        elif col == ColNum.STEP_NOTE.value:
            step.notetext = value

    def find_step_index(self, row):
        for index, in_step_index in enumerate(self.step_list):
            if in_step_index[0] == row:
                return index
        return None

    def find_step_number(self, row):
        num = self._safe_step_number(row)
        return num

    def find_box_index(self, row):
        for index, in_box_index in enumerate(self.box_list):
            if in_box_index[0] == row:
                return index
        return None

    def closeEvent(self, event: QCloseEvent) -> None:
        save_data_list = []
        insert_row = 1

        filtered_steps = [
            step for _, step in self.step_list
            if hasattr(step, "type_") and not (step.type_ == procedure.Type.Sub and not step.condition)
        ]

        print(f"Filtered steps: {len(filtered_steps)}")

        for step in filtered_steps:
            step_data = []
            step_dict = step.__dict__.copy()

            step_type = step_dict.pop("type_")
            step_name = getattr(step_type, "name", str(step_type))

            if step_name == "Sub":
                step_data.extend([None, step_name])
            else:
                step_data.extend([insert_row, step_name])
                insert_row += 1

            if step_name == "EIS":
                try:
                    eis_match = next((e for e in self.eis_parameters if e[0] == insert_row - 1), None)
                    eis_path = eis_match[-1] if eis_match else None
                except Exception as e:
                    print("EIS 매칭 중 오류:", e)
                    eis_path = None
                step_data.append(eis_path)

            for val in step_dict.values():
                if isinstance(val, (float, int, str, list)) or val is None:
                    step_data.append(val)
                else:
                    step_data.append(getattr(val, "name", str(val)))
            if step_name == "EIS":
                del (step_data[4])
            if step_data[1] == "Sub":
                step_data[11] = None
            if step_name != "EIS":
                step_data.append(None)
            save_data_list.append(step_data)

        if filtered_steps and getattr(filtered_steps[-1].type_, "name", "") != "End":
            save_data_list.append([insert_row, "End"] + [None] * 10)

        print("==============")
        print("✅ 변환 완료:", len(save_data_list), "rows")
        print("save data : ", save_data_list)
        print("==============")

        if self.file is None:
            self.file = (QFileDialog.getSaveFileName(self, "Save File", '', 'xlsx(*.xlsx)'))[0]

        if not self.eis_parameters:
            save_esidata_list = [[None] * 8]
        else:
            save_esidata_list = [row[:-1] + [None, None] for row in self.eis_parameters]

        print("EIS 저장 데이터:", save_esidata_list)
        file.save_full_excel(self.file, save_data_list, save_esidata_list)

        return super().closeEvent(event)

    def add_or_update_eis_parameter(self, new_row: list):
        if not new_row or len(new_row) < 2:
            print("[EIS] 잘못된 데이터 형식:", new_row)
            return

        row_index = new_row[0]

        for i, existing_row in enumerate(self.eis_parameters):
            if existing_row[0] == self.__row:
                new_row.insert(0, self.__row)
                self.eis_parameters[i] = new_row
                print(f"[EIS] 행 {row_index} 데이터 갱신 완료")
                break
        else:
            new_row.insert(0, self.__row)
            self.eis_parameters.append(new_row)
            print(f"[EIS] 행 {row_index} 데이터 추가 완료")

    def add_eis_parameters(self, params):
        self.add_or_update_eis_parameter(params[0])
        print(self.eis_parameters)
        self.generate_value_cell(self.__row, 2, str(params[0][-1]))

    def insert_step_above(self, any_row: int, *, do_reset: bool = True):
        total_rows = self.tableWidget.rowCount()
        if total_rows == 0:
            insert_pos = 0
        else:
            main_row = any_row
            while main_row >= 0:
                item = self.tableWidget.item(main_row, ColNum.STEP.value)
                if item and item.text().strip().isdigit():
                    break
                main_row -= 1
            if main_row < 0:
                main_row = 0
            insert_pos = main_row

        for s in self.step_list:
            if s[0] >= insert_pos:
                s[0] += 2
        for b in self.box_list:
            if b[0] >= insert_pos:
                b[0] += 2
        for s in self.step_list:
            step_obj = s[1]
            if hasattr(step_obj, "depend_row") and step_obj.depend_row is not None:
                if step_obj.depend_row >= insert_pos:
                    step_obj.depend_row += 2

        self.tableWidget.insertRow(insert_pos)
        self.new_step(insert_pos)
        for col in (
            ColNum.TYPE.value,
            ColNum.MODE.value,
            ColNum.END_TYPE.value,
            ColNum.OP.value,
            ColNum.REPORT_TYPE.value,
        ):
            self.new_cell(insert_pos, col, method_name=self.generate_combo_box)
        for col in (
            ColNum.MODE_VALUE.value,
            ColNum.OP_VALUE.value,
            ColNum.GOTO.value,
            ColNum.REPORT_VALUE.value,
            ColNum.STEP_NOTE.value,
        ):
            self.new_cell(insert_pos, col, method_name=self.generate_value_cell)

        blank_row = insert_pos + 1
        self.tableWidget.insertRow(blank_row)
        self.new_cell(blank_row, ColNum.STEP.value, text="")

        if do_reset:
            self.reset_step_index()
            self.reset_goto()

        return insert_pos

    def delete_step_block(self, any_row: int):
        total_rows = self.tableWidget.rowCount()
        if total_rows == 0:
            return

        main_row = any_row
        while main_row >= 0:
            item = self.tableWidget.item(main_row, ColNum.STEP.value)
            if item and item.text().strip().isdigit():
                break
            main_row -= 1
        if main_row < 0:
            return

        main_rows = []
        for r in range(total_rows):
            it = self.tableWidget.item(r, ColNum.STEP.value)
            if it and it.text().strip().isdigit():
                main_rows.append(r)
        if not main_rows:
            return

        if main_row == main_rows[-1]:
            print("[INFO] 마지막 메인 스텝은 삭제하지 않습니다.")
            return

        block_end = main_row
        for r in range(main_row + 1, total_rows):
            step_item = self.tableWidget.item(r, ColNum.STEP.value)
            if (not step_item) or (step_item.text().strip() == ""):
                block_end = r
                nxt = r + 1
                if nxt >= total_rows:
                    break
                nxt_item = self.tableWidget.item(nxt, ColNum.STEP.value)
                if nxt_item and nxt_item.text().strip().isdigit():
                    break
            else:
                break

        delete_rows = list(range(main_row, block_end + 1))
        deleted_count = len(delete_rows)

        for r in reversed(delete_rows):
            self.tableWidget.removeRow(r)

        self.step_list = [s for s in self.step_list if not (main_row <= s[0] <= block_end)]
        self.box_list = [b for b in self.box_list if not (main_row <= b[0] <= block_end)]

        for s in self.step_list:
            if s[0] > block_end:
                s[0] -= deleted_count
        for b in self.box_list:
            if b[0] > block_end:
                b[0] -= deleted_count
        for s in self.step_list:
            step_obj = s[1]
            if hasattr(step_obj, "depend_row") and step_obj.depend_row is not None:
                if step_obj.depend_row > block_end:
                    step_obj.depend_row -= deleted_count

        self.reset_step_index()
        self.reset_goto()

    def __copy_step(self, any_row: int):
        main_row, last_row = self.__find_block_bounds(any_row)
        if main_row is None or last_row is None:
            print("[COPY] 대상 블록을 찾지 못했습니다.")
            return

        block = self.__collect_block_steps(main_row, last_row)
        if not block:
            print("[COPY] 복사할 Step이 없습니다.")
            return

        self.paste_list = copy.deepcopy(block)
        print(f"[COPY] rows {main_row}..{last_row} ({len(self.paste_list)} steps) 복사 완료.")

    def paste_step_above(self, target_row: int, copied_step: "procedure.Step"):
        self.tableWidget.blockSignals(True)
        try:
            insert_pos = self.insert_step_above(target_row, do_reset=False)

            self.new_cell(insert_pos, ColNum.TYPE.value, method_name=self.generate_combo_box)
            self.__change_combo_box(
                insert_pos, ColNum.TYPE.value,
                copied_step.type_.name, self.change_type, isLoad=1
            )

            if getattr(copied_step, "mode", None) is not None:
                self.new_cell(insert_pos, ColNum.MODE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(
                    insert_pos, ColNum.MODE.value,
                    copied_step.mode.name, self.change_mode, isLoad=1
                )

            if getattr(copied_step, "condition", None) is not None:
                self.new_cell(insert_pos, ColNum.END_TYPE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(
                    insert_pos, ColNum.END_TYPE.value,
                    copied_step.condition.name, self.change_condition, isLoad=1
                )

            if getattr(copied_step, "operator", None) is not None:
                self.new_cell(insert_pos, ColNum.OP.value, method_name=self.generate_combo_box)
                self.__change_combo_box(
                    insert_pos, ColNum.OP.value,
                    str(copied_step.operator), self.change_operator, isLoad=1
                )

            if getattr(copied_step, "report", None) is not None:
                self.new_cell(insert_pos, ColNum.REPORT_TYPE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(
                    insert_pos, ColNum.REPORT_TYPE.value,
                    copied_step.report.name, self.change_report, isLoad=1
                )

            if getattr(copied_step, "mode_val", None) is not None:
                self.__change_value(insert_pos, ColNum.MODE_VALUE.value, copied_step.mode_val)
            if getattr(copied_step, "condition_val", None) is not None:
                self.__change_value(insert_pos, ColNum.OP_VALUE.value, copied_step.condition_val)
            if getattr(copied_step, "report_val", None) is not None:
                self.__change_value(insert_pos, ColNum.REPORT_VALUE.value, copied_step.report_val)
            if getattr(copied_step, "notetext", None) is not None:
                self.__change_value(insert_pos, ColNum.STEP_NOTE.value, copied_step.notetext)

            self.step_list.sort(key=lambda x: x[0])
            self.reset_step_index()
            self.reset_goto()
        finally:
            self.tableWidget.blockSignals(False)

    @Slot(QPoint)
    def __context_menu(self, pos):
        menu = QMenu()
        copy_action = menu.addAction("Copy Step")
        paste_action = menu.addAction("Paste Step")
        menu.addSeparator()
        insert_action = menu.addAction("Insert Step")
        delete_action = menu.addAction("Delete Step")

        # (선택) Undo/Redo 메뉴 추가 가능
        undo_action = menu.addAction("Undo")
        redo_action = menu.addAction("Redo")

        action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        row = self.tableWidget.indexAt(pos).row()

        if action == copy_action:
            self.__copy_step(row)
        elif action == paste_action:
            if self.paste_list:
                cmd = PasteStepCommand(self, row)
                self.execute_command(cmd)
        elif action == insert_action:
            cmd = InsertStepCommand(self, row)
            self.execute_command(cmd)
        elif action == delete_action:
            cmd = DeleteStepBlockCommand(self, row)
            self.execute_command(cmd)
        elif action == undo_action:
            self.undo()
        elif action == redo_action:
            self.redo()

    def __find_block_bounds(self, any_row: int):
        total = self.tableWidget.rowCount()
        if total == 0:
            return None, None

        main_row = any_row
        while main_row >= 0:
            it = self.tableWidget.item(main_row, ColNum.STEP.value)
            if it and it.text().strip().isdigit():
                break
            main_row -= 1
        if main_row < 0:
            return None, None

        last_step_row = main_row
        for r in range(main_row + 1, total):
            it = self.tableWidget.item(r, ColNum.STEP.value)
            if it and it.text().strip().isdigit():
                break
            if not it or it.text().strip() == "":
                break
            last_step_row = r

        return main_row, last_step_row

    def __collect_block_steps(self, main_row: int, last_row: int):
        result = []
        rows = set(range(main_row, last_row + 1))
        for row, step in self.step_list:
            if row in rows and hasattr(step, "type_"):
                result.append((row, copy.deepcopy(step)))
        return result

    def _cell_text(self, row, col):
        item = self.tableWidget.item(row, col)
        return item.text().strip() if item else ""

    def _is_int_text(self, s: str) -> bool:
        return s.isdigit()

    def _safe_step_number(self, row):
        txt = self._cell_text(row, ColNum.STEP.value)
        return int(txt) if self._is_int_text(txt) else None

    def _set_step_text(self, row, text):
        self.new_cell(row, ColNum.STEP.value, text=str(text) if text is not None else "")