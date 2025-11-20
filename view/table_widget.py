# pylint: disable = no-name-in-module

# from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent, QContextMenuEvent, QAction, QCursor
from PySide6.QtWidgets import ( QWidget, QFileDialog, QTableWidgetItem, QMenu, QAbstractItemView,
    QComboBox, QLineEdit)
from PySide6.QtCore import Qt, QEvent, Slot, QPoint

from view import ui_BuildTestTableWidget
from model import procedure
from model import procedure_copy
from model import file
import pandas as pd
import pickle
import enum
import copy
import openpyxl

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

class Box_method_item():
    type:None
    mode:None
    endtype:None
    operator:None
    retype:None

#====================Table Window==========================#
class TableWidget(QWidget, ui_BuildTestTableWidget.Ui_BuildTestTableWidget):
    def __init__(self, status_widgets=None, file_name=None) -> None:
        super().__init__()
        self.setupUi(self)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.__row = 0
        self.__col = 0
        set_col = 0
        self.file = file_name
        self.status_widgets = status_widgets
        self.init_rest =["Rest","Step_time","=",3,"2","Step_time",1]
        self.step_list = []
        self.box_list=[]
        self.load_step_list =[]
        self.eis_parameters = []
        self.sub_step_cnt = 0

        self.paste_list = []
        self.paste_box_list = []

        #Set signal Event
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.__context_menu)
        self.tableWidget.cellClicked.connect(self.__click_cell)
        self.tableWidget.cellChanged.connect(self.__changed_cell)
        try:
            if self.file is None:
                self.new_step(0)
                self.new_step(2)
                self.new_cell(2, 1, method_name=self.generate_combo_box)

                self.reset_step_index()
                for col, item in enumerate(self.init_rest):
                    if col < ColNum.TYPE.value:
                        set_col = col + 1
                    else:
                        set_col = col + 3
                    if set_col in (ColNum.TYPE.value, ColNum.END_TYPE.value, ColNum.OP.value, ColNum.REPORT_TYPE.value):
                        self.new_cell(0, set_col, text=item, method_name=self.generate_combo_box)
                        self.__change_combo_box(0, set_col, str(item), self.find_method_box(set_col))
                    else:
                        self.new_cell(0, set_col, text=item, method_name=self.generate_value_cell)
                        self.__change_value(0, set_col, item)

            else:
                # 1) 파일 로드
                self.load_step_list, temp_eis = file.load_full_excel(file_name)

                # 2) 헤더 제거
                if self.load_step_list:
                    del self.load_step_list[0]
                if temp_eis:
                    del temp_eis[0]

                # 3) EIS 경로 매칭
                for eis_item in temp_eis:
                    eis_path = self.search_eis_path(eis_item[0], self.load_step_list)
                    eis_item.append(eis_path)
                    self.eis_parameters.append(eis_item)

                # 4) 메인 스텝 시작 인덱스 수집
                load_step_package = []
                package_main_index = []

                for idx, row in enumerate(self.load_step_list):
                    if row:
                        del row[-1]
                    if row and not pd.isna(row[0]):
                        package_main_index.append(idx)

                if not package_main_index:
                    package_main_index = [0]
                package_main_index.append(len(self.load_step_list))

                # 5) 메인 스텝 단위로 패키징
                for i, start in enumerate(package_main_index[:-1]):
                    end = package_main_index[i + 1]
                    load_step_package.append(self.load_step_list[start:end])

                # 6) 테이블에 로드 (수정된 load_step 사용)
                set_row = 0
                for pkg in load_step_package:
                    self.load_step(pkg, set_row)   # ← 여기서 수정된 버전이 호출되어야 함
                    set_row += (len(pkg) + 1)

        except Exception as e:
            print("load error:", e)

    def load_step(self, load_step_list, set_row):
        import math
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
                        step_obj.depend_row = prev_main_row  # 임시로 row 번호 저장 (나중에 매핑됨)
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

        # 전체 로드 후 Step 인덱스 정리
        self.reset_step_index()

        # 🔹 Sub Step의 depend_row를 step_list 기준으로 다시 매핑
        for step in self.step_list:
            if getattr(step[1], "type_", None) == procedure.Type.Sub:
                # depend_row(행번호) → step_list 내 인덱스로 변환
                dep_idx = self.find_step_index(step[1].depend_row)
                if dep_idx is not None:
                    step[1].depend_row = self.step_list[dep_idx][0]
                else:
                    step[1].depend_row = 0  # 안전장치




    def search_eis_path(self, step_num: int, step_list: list[list]) -> str | None:
        """
        step_list에서 해당 step_num의 EIS 경로를 찾아 반환.
        예: step_list = [[1, 'Rest', ...], [2, 'EIS', 'C:/path.xlsx', ...]]
        """
        for step in step_list:
            if not step or len(step) < 3:
                continue
            if str(step[0]) == str(step_num) and str(step[1]).upper() == "EIS":
                return step[2]  # 3번째 컬럼이 경로
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

    """------ Signal Event ------"""

    def __click_cell(self, row, col):
        ### Select Cell Signal processing
        self.tableWidget.removeCellWidget(self.__row, self.__col)

        if not self.box_list[0]: # procedure box method list check
            print("null box_list_item list create" )
        else:
            for row_method in self.box_list: # box list exists check
                if row_method[0] == row:
                    now_box_method = row_method[1]

        cell_item_call = self.tableWidget.item(row,col)

        if cell_item_call and cell_item_call.function != None : #combo Box create
            cell_item_call.function(row, col, self.tableWidget.item(row, col).text(), box_method = now_box_method)

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
                parent_window = self.window()  # MainWindow 객체
                if hasattr(parent_window, "open_eis_window"):
                    parent_window.open_eis_window(find_step_number)
                #self.new_cell(row, ColNum.OP_VALUE, method_name = self.generate_value_cell)
                self.change_condition(row, " ", 0)

            if col == ColNum.END_TYPE.value:
                if stepType == procedure.Type.Sub:
                    if not self.find_step_obj(row + 1):
                        self.sub_step(row) #서브 조건 시작

                elif stepType == procedure.Type.Charge or stepType == procedure.Type.Discharge:
                    self.sub_step(row,1)

    def find_step_obj(self, row):
        find_step = None
        for i in self.step_list:
            if i[0] == row:
                find_step = i[1]
        return find_step
    
    def reset_goto(self):
        """
        goto 재계산:
        - Main: 자신의 Step 번호 + 1
        - Sub : 부모 Main의 goto 그대로
        """
        # 1) Main 먼저 계산
        for row, step in self.step_list:
            if not hasattr(step, "type_") or step.type_ == procedure.Type.Sub:
                continue

            step_no = self._safe_step_number(row)
            if step_no is None:
                # 번호 셀이 비어 있으면 스킵 (구분행 등)
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

            parent_row, parent_step = self.step_list[parent_idx]
            parent_goto_txt = self._cell_text(parent_row, ColNum.GOTO.value)
            if parent_goto_txt:
                self.__change_value(row, ColNum.GOTO.value, parent_goto_txt)

    # def reset_goto(self):
    #     """
    #     각 Step의 goto 값을 재설정한다.
    #     - Main Step은 자신의 다음 Step 번호로 goto 지정
    #     - Sub Step은 의존하는 Main Step의 goto를 따라감
    #     """
    #     # 🔹 1차 루프: Main Step / Cycle Step 처리
    #     for step in self.step_list:
    #         if not hasattr(step[1], "type_"):
    #             continue

    #         # 조건문 Step (Main)
    #         if getattr(step[1], "condition", False):
    #             if step[1].type_ != procedure.Type.Sub:
    #                 try:
    #                     goto = int(self.tableWidget.item(step[0], 0).text()) + 1
    #                     self.__change_value(step[0], 7, goto)
    #                 except Exception:
    #                     continue

    #         # Cycle Start / End Step 처리
    #         elif step[1].type_ in (procedure.Type.Cycle_start, procedure.Type.Cycle_end):
    #             try:
    #                 goto = int(self.tableWidget.item(step[0], 0).text()) + 1
    #                 self.__change_value(step[0], 7, goto)
    #             except Exception:
    #                 continue

    #     # 🔹 2차 루프: Sub Step 처리
    #     for step in self.step_list:
    #         if not hasattr(step[1], "type_"):
    #             continue

    #         if getattr(step[1], "condition", False) and step[1].type_ == procedure.Type.Sub:
    #             depend_row = getattr(step[1], "depend_row", None)
    #             if depend_row is None:
    #                 # Sub Step이 의존할 Main Step이 없으면 스킵
    #                 continue

    #             # depend_row가 step 객체가 아니라면 step_list에서 인덱스를 찾아야 함
    #             dep_idx = self.find_step_index(depend_row)

    #             if dep_idx is None:
    #                 # 유효한 인덱스가 아니면 스킵
    #                 continue

    #             try:
    #                 goto_value = getattr(self.step_list[dep_idx][1], "goto", None)
    #                 if goto_value is None:
    #                     continue
    #                 goto = int(goto_value)
    #                 self.__change_value(step[0], 7, goto)
    #             except Exception:
    #                 continue

    # def reset_goto(self):
    #     for step in self.step_list:
    #         if step[1].__dict__.__contains__("type_"):
    #             if step[1].condition:
    #                 if step[1].type_ != procedure.Type.Sub:
    #                     goto = int(self.tableWidget.item(step[0],0).text()) + 1
    #                     self.__change_value(step[0],7,goto)
    #             elif step[1].type_ == procedure.Type.Cycle_start:
    #                 goto = int(self.tableWidget.item(step[0],0).text()) + 1
    #                 self.__change_value(step[0],7,goto)
    #             elif step[1].type_ == procedure.Type.Cycle_end :
    #                 goto = int(self.tableWidget.item(step[0],0).text()) + 1
    #                 self.__change_value(step[0],7,goto)

    #     for step in self.step_list:
    #         if step[1].__dict__.__contains__("type_"):
    #             if step[1].condition:
    #                 if step[1].type_ == procedure.Type.Sub:
    #                     goto = int( self.step_list[self.find_step_index(step[1].depend_row)][1].goto )
    #                     self.__change_value(step[0],7,goto)

    def new_step(self, row):
        ### Create Step, ComboBox
        # row : current position row
        # col : current position column
        new_step = [row, procedure.Step()]
        new_box_method = [row, Box_method_item()]

        self.step_list.append(new_step)
        self.box_list.append(new_box_method)

        self.step_list.sort(key=lambda x:x[0])
        self.box_list.sort(key=lambda x:x[0])

    def update_row(self, Current_row):
        ### Current Row ~ Last Row Update
        for row_index in range(Current_row, len(self.step_list)):
            self.step_list[row_index][0] += 1
            self.box_list[row_index][0] += 1

    def sub_step(self, row, new_sub=None):
        sub_step_row = row + 1
        self.new_step(sub_step_row)

        update_start_index = self.find_step_index(sub_step_row)
        self.update_row(update_start_index + 1) # +1 : Sub 제외 ++

        step = self.step_list[update_start_index][1]
        step.type_ = ( procedure.Type.str_to_enum("Sub")  )

        if new_sub:
            step.depend_row = row
        else:
            step.depend_row = self.find_step_obj(row).depend_row

        for i in range(update_start_index + 1, len(self.step_list)):
            if self.step_list[i][1].__dict__.__contains__("type_"):
                if self.step_list[i][1].type_ == procedure.Type.Sub:
                    self.step_list[i][1].depend_row +=1

        self.tableWidget.insertRow(sub_step_row)
        sub_combo_list = [4,8]
        for coombo_col in sub_combo_list:
            self.new_cell(sub_step_row, coombo_col, method_name = self.generate_combo_box)

        self.sub_step_cnt +=1

    def new_cell(self, row, col, text: str=" ", method_name=None):
        ### Set cell Text And ComboBox method (Set row is absolute)
        # row : current position row
        # col : current position column
        # text : current position Text item
        # method_name : current cell method name ( is ComboBox method )
        if pd.isna(text):
            text = " "
        cell = TableItem(text)
        cell.function = method_name
        self.tableWidget.removeCellWidget(row, col)
        self.tableWidget.setItem(row, col, cell)


    def generate_combo_box(self, row, col, text=" ", box_method=None, isLoad=None):
        """
        지정된 셀(row, col)에 QComboBox를 생성하고 Step 데이터와 연결합니다.
        """
        # ───────────────────────────────────────────────
        # ① Step 객체 및 인덱스 확인
        # ───────────────────────────────────────────────
        step_index = self.find_step_index(row)
        if step_index is None or step_index >= len(self.step_list):
            return

        step = self.step_list[step_index][1]

        # Step 객체에 type_ 속성이 없을 경우 기본값 지정 (AttributeError 방지)
        if not hasattr(step, "type_"):
            step.type_ = None

        # 콤보박스 생성
        combo = QComboBox()
        combo.setPlaceholderText(text)
        combo.setCurrentIndex(-1)

        # ───────────────────────────────────────────────
        # ② Sub Step의 부모 Type 참조
        # ───────────────────────────────────────────────
        type_condition = None
        if getattr(step, "type_", None) == procedure.Type.Sub:
            depend_row = getattr(step, "depend_row", None)
            if depend_row is not None:
                parent_idx = self.find_step_index(depend_row)
                if parent_idx is not None:
                    parent_step = self.step_list[parent_idx][1]
                    if hasattr(parent_step, "type_"):
                        type_condition = parent_step.type_

        # ───────────────────────────────────────────────
        # ③ 컬럼에 따른 콤보박스 항목 및 함수 설정
        # ───────────────────────────────────────────────
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

        # ───────────────────────────────────────────────
        # ④ 유효한 콤보 설정이 없으면 종료
        # ───────────────────────────────────────────────
        if not col_config or not col_config.get("items"):
            return

        items = col_config["items"]
        change_func = col_config["func"]
        attr_name = col_config["attr"]

        # 중복 항목 방지 (현재 표시 중인 텍스트 제거)
        if text in items:
            items.remove(text)

        # ───────────────────────────────────────────────
        # ⑤ 항목 추가 및 시그널 연결
        # ───────────────────────────────────────────────
        combo.addItems(items)
        combo.currentTextChanged.connect(
            lambda value: self.__change_combo_box(row, col, value, change_func, isLoad)
        )

        # box_method가 있을 경우 해당 속성으로 보관
        if box_method:
            setattr(box_method, attr_name, combo)

        # ───────────────────────────────────────────────
        # ⑥ 테이블에 콤보박스 적용
        # ───────────────────────────────────────────────
        self.tableWidget.setCellWidget(row, col, combo)

    def __change_combo_box(self,row,col,text,change_func,isLoad=None):
        change_func(row,text,isLoad)
        if col==ColNum.TYPE.value and text=="Sub": return
        item=self.tableWidget.item(row,col)
        (item.setText(text) if item else self.new_cell(row,col,text=text,method_name=self.generate_value_cell))


    def change_type(self, row, text, isLoad=None):
        """
        step.type_ 변경 시 테이블 셀 구성 및 콤보박스 재생성 처리
        """
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.type_ = procedure.Type.str_to_enum(text)

        if isLoad:
            return
        self.reset_val(row)

        # 공통적으로 재생성되는 셀들 (콤보 + 값)
        combo_cols = [ColNum.MODE.value, ColNum.END_TYPE.value, ColNum.REPORT_TYPE.value, ColNum.OP.value]
        value_cols = [ColNum.MODE_VALUE.value, ColNum.OP_VALUE.value, ColNum.REPORT_VALUE.value]

        for col in combo_cols:
            self.new_cell(row, col, method_name=self.generate_combo_box)
        for col in value_cols:
            self.new_cell(row, col, method_name=self.generate_value_cell)

        # Goto 및 StepNote 셀 공통 생성
        self.new_cell(row, ColNum.GOTO.value, method_name=self.generate_value_cell)
        self.new_cell(row, ColNum.STEP_NOTE.value, method_name=self.generate_value_cell)

        if step.type_ == procedure.Type.Cycle_start:
            # 특별한 기본값 없음
            self.reset_goto()

        elif step.type_ == procedure.Type.Cycle_end:
            # Cycle_end 전용 기본값
            self.__set_combo_and_apply(row, ColNum.END_TYPE.value, "Cycle_Count", self.change_condition)
            self.__set_combo_and_apply(row, ColNum.OP.value, "=", self.change_operator)
            self.__set_combo_and_apply(row, ColNum.REPORT_TYPE.value, "Step_time", self.change_report)
            self.reset_goto()

        elif step.type_ == procedure.Type.DCIR:
            # DCIR 전용 기본값
            self.__set_combo_and_apply(row, ColNum.MODE.value, "Current", self.change_mode)
            self.__set_combo_and_apply(row, ColNum.REPORT_TYPE.value, "Step_time", self.change_report)
            self.__set_combo_and_apply(row, ColNum.END_TYPE.value, "Step_time", self.change_condition)
            self.__set_combo_and_apply(row, ColNum.OP.value, "=", self.change_operator)
            # 빈 값 셀 생성
            self.new_cell(row, ColNum.OP_VALUE.value, text=" ", method_name=self.generate_value_cell)
            self.new_cell(row, ColNum.REPORT_VALUE.value, text=" ", method_name=self.generate_value_cell)
            self.new_cell(row, ColNum.STEP_NOTE.value, text=" ", method_name=self.generate_value_cell)

        # ──────────────────────────────
        # ③ 공통 후처리: 다음 Step 존재 여부 확인
        # ──────────────────────────────
        if (
            self.step_list[-1][0] == row
            and text != "End"
            and self.step_list[-1][1].__dict__.__contains__("type_")
        ):
            self.new_step(self.step_list[-1][0] + 2)
            self.new_cell(self.step_list[-1][0], ColNum.TYPE.value, method_name=self.generate_combo_box)
            self.reset_step_index()

    def __set_combo_and_apply(self, row, col, value, change_func):
        """콤보박스 셀을 생성하고 즉시 값 반영"""
        self.new_cell(row, col, text=value, method_name=self.generate_combo_box)
        self.__change_combo_box(row, col, str(value), self.find_method_box(col))
        change_func(row, value, 0)


    def change_mode(self, row, text, isLoad =None): #No sub Step
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.mode = (procedure.Mode.str_to_enum(text))
        if not isLoad:
            self.new_cell(row, 3, method_name = self.generate_value_cell)

    def reset_val(self, row):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.mode = None
        step.mode_val = None
        step.condition = None
        step.operator = None
        step.condition_val =None
        step.report = None
        step.report_val =None
        step.notetext = None

    def change_condition(self, row, text, isLoad =None): #Sub step
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.condition = (procedure.Condition.str_to_enum(text))

        if isLoad:
            pass
        else:
            if self.tableWidget.item(row,0):
                step_item = self.tableWidget.item(row,0).text()
                if step_item != ' ':
                    insert_goto = int(step_item) + 1
            if step.type_ == procedure.Type.Sub:
                find_root_step = self.step_list[self.find_step_index(step.depend_row)][1]
                step.goto = find_root_step.goto
                insert_goto = step.goto
            else: #not sub
                step.goto = (str(insert_goto))
            self.new_cell(row, 5, method_name=self.generate_combo_box) # opreator
            self.new_cell(row, 6, method_name = self.generate_value_cell)
            self.new_cell(row, 7, text=str(insert_goto), method_name = self.generate_value_cell)

    def change_operator(self, row, text, isLoad =None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.operator = text
        print("여기 op 는 : ",step_index,row)

    def change_report(self, row, text, isLoad =None):
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]
        step.report = procedure.Report.str_to_enum(text)
        if not isLoad:
            self.new_cell(row, 9, method_name = self.generate_value_cell)

    def generate_value_cell(self, row, col, text=" ", box_method= None):
        value_cell = QLineEdit(text)
        value_cell.textChanged.connect(
            lambda _text: self.__change_value(row, col, _text)
        )
        self.tableWidget.setCellWidget(row, col, value_cell)

    def reset_step_index(self):
        """
        UI 테이블의 Step(번호) 칼럼만 일관되게 다시 써준다.
        - Main step: 1,2,3,... 번호 부여
        - Sub step: 빈칸("") 유지
        """
        seq = 1
        # step_list는 실제 step들만 들어 있으므로 이 순서대로 테이블 셀 갱신
        for row, step in sorted(self.step_list, key=lambda x: x[0]):
            if hasattr(step, "type_") and step.type_ == procedure.Type.Sub:
                # Sub는 Step 칸 비움
                self._set_step_text(row, "")
            else:
                self._set_step_text(row, seq)
                seq += 1


    def __change_value(self, row, col, text):
        try:
            print(row, col, text)
            if col == ColNum.STEP.value: # 0 column is unique seq(Step Number)
                return
            elif col == ColNum.GOTO.value: # goto
                value = text
            elif col == ColNum.STEP_NOTE.value:
                value = str(text)
            else:
                value = float(text)
            print(self.tableWidget.item(row,col))
            self.tableWidget.item(row, col).setText(str(value))

        except Exception as e:
            #self.textBrowser.setText("value type error{}".format(text)) #텍스트 LOG판 여기
            print(e)
            return
        #Step value set
        step_index = self.find_step_index(row)
        step = self.step_list[step_index][1]

        if col == ColNum.MODE_VALUE.value:
            step.mode_val = value
        elif col == ColNum.OP_VALUE.value:
            step.condition_val  = value
        elif col == ColNum.GOTO.value:
            step.goto = value
        elif col == ColNum.REPORT_VALUE.value:
            step.report_val = value
        elif col == ColNum.STEP_NOTE.value:
            step.notetext = value

    def find_step_index(self, row):
        ### Find Main Step index
        # row : current position row
        re_index = None
        for index, in_step_index in enumerate(self.step_list):
            if in_step_index[0] == row:
                re_index = index
                return re_index

    def find_step_number(self, row):
        num = self._safe_step_number(row)
        return num  # 없으면 None


    def find_box_index(self, row):
        ### Find Step List index
        # row : current position row
        re_index = None
        for index, in_box_index in enumerate(self.box_list):
            if in_box_index[0] == row:
                re_index = index
                return re_index

    def closeEvent(self, event: QCloseEvent) -> None:
        save_data_list = []
        insert_row = 1

        # 불필요한 step 필터링
        filtered_steps = [
            step for _, step in self.step_list
            if hasattr(step, "type_") and not (step.type_ == procedure.Type.Sub and not step.condition)
        ]

        print(f"Filtered steps: {len(filtered_steps)}")

        # Step별 데이터 파싱
        for step in filtered_steps:
            step_data = []
            step_dict = step.__dict__.copy()

            # 타입 정보 추출
            step_type = step_dict.pop("type_")
            step_name = getattr(step_type, "name", str(step_type))

            # --- Step 번호 및 타입 ---
            if step_name == "Sub":
                # Sub 타입: Step 번호 없음
                step_data.extend([None, step_name])
            else:
                # 일반 Step 및 EIS 타입: Step 번호 포함
                step_data.extend([insert_row, step_name])
                insert_row += 1

            # --- EIS 타입일 경우 EIS 파라미터 연결 ---
            if step_name == "EIS":
                try:
                    # 동일 Step 번호에 해당하는 EIS 데이터 찾기
                    eis_match = next((e for e in self.eis_parameters if e[0] == insert_row - 1), None)
                    eis_path = eis_match[-1] if eis_match else None
                except Exception as e:
                    print("EIS 매칭 중 오류:", e)
                    eis_path = None
                step_data.append(eis_path)

            # --- 나머지 속성 값 추가 ---
            for val in step_dict.values():
                if isinstance(val, (float, int, str, list)) or val is None:
                    step_data.append(val)
                else:
                    step_data.append(getattr(val, "name", str(val)))
            if step_name == "EIS":
                del(step_data[4])  # EIS 경로 중복 제거
            # --- 마지막 한 칸은 공백(None) ---
            if(step_data[1] == "Sub"):
                step_data[11] = None
            if step_name != "EIS":
                step_data.append(None)
            save_data_list.append(step_data)

        # 3️⃣ 마지막 Step이 End가 아니면 자동 추가
        if filtered_steps and getattr(filtered_steps[-1].type_, "name", "") != "End":
            save_data_list.append([insert_row, "End"] + [None] * 10)

        print("==============")
        print("✅ 변환 완료:", len(save_data_list), "rows")
        print("save data : ", save_data_list)
        print("==============")

        # 4️⃣ 파일 경로 지정
        if self.file is None:
            self.file = (QFileDialog.getSaveFileName(self, "Save File", '', 'xlsx(*.xlsx)'))[0]

        # 5️⃣ EIS 데이터 저장용 변환
        if not self.eis_parameters:
            save_esidata_list = [[None]*8]
            print("")
        else:
            # 경로를 제거하고 None 2개를 추가 (엑셀 구분용)
            save_esidata_list = [row[:-1] + [None, None] for row in self.eis_parameters]
            print("")

        print("EIS 저장 데이터:", save_esidata_list)

        # 6️⃣ 통합 저장
        file.save_full_excel(self.file, save_data_list, save_esidata_list)

        return super().closeEvent(event)

    def add_or_update_eis_parameter(self, new_row: list):
        """
        EIS 설정 데이터를 self.eis_parameters에 추가 또는 갱신.
        new_row 형식: [row_index, mode, start_f, stop_f, amp, point_num, file_path]
        """
        if not new_row or len(new_row) < 2:
            print("[EIS] 잘못된 데이터 형식:", new_row)
            return

        row_index = new_row[0]

        # 이미 동일한 row_index가 존재하는지 검사
        for i, existing_row in enumerate(self.eis_parameters):
            if existing_row[0] == self.__row:
                # 기존 행 업데이트
                new_row.insert(0, self.__row)  # row_index 추가
                self.eis_parameters[i] = new_row
                print(f"[EIS] 행 {row_index} 데이터 갱신 완료")
                break
        else:
            # 존재하지 않으면 새로 추가
            new_row.insert(0, self.__row)  # row_index 추가
            self.eis_parameters.append(new_row)
            print(f"[EIS] 행 {row_index} 데이터 추가 완료")

    def add_eis_parameters(self, params):
        """MainWindow에서 받은 EIS 설정값을 리스트에 저장"""
        self.add_or_update_eis_parameter(params[0])
        print(self.eis_parameters)
        self.generate_value_cell(self.__row, 2, str(params[0][-1]))
        #self.tableWidget.setItem(int(params[0][0]), 2, str(params[0][-1]))

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

        # 인덱스 보정
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

        # 메인 행 + 구분행 삽입
        self.tableWidget.insertRow(insert_pos)
        self.new_step(insert_pos)
        for col in (ColNum.TYPE.value, ColNum.MODE.value, ColNum.END_TYPE.value,
                    ColNum.OP.value, ColNum.REPORT_TYPE.value):
            self.new_cell(insert_pos, col, method_name=self.generate_combo_box)
        for col in (ColNum.MODE_VALUE.value, ColNum.OP_VALUE.value,
                    ColNum.GOTO.value, ColNum.REPORT_VALUE.value, ColNum.STEP_NOTE.value):
            self.new_cell(insert_pos, col, method_name=self.generate_value_cell)

        blank_row = insert_pos + 1
        self.tableWidget.insertRow(blank_row)
        self.new_cell(blank_row, ColNum.STEP.value, text="")

        if do_reset:
            self.reset_step_index()
            self.reset_goto()

        return insert_pos  # 새 메인 스텝이 들어간 실제 위치를 반환


    def delete_step_block(self, any_row: int):
        """
        현재 위치(any_row)가 속한 '메인 스텝 블록(메인 + 모든 Sub + 구분자)'을 삭제한다.
        단, '마지막 메인 스텝'이면 삭제하지 않고 바로 반환한다.
        (구분자 row는 table에서만 제거, 내부 리스트는 메인/서브만 정리)
        """
        total_rows = self.tableWidget.rowCount()
        if total_rows == 0:
            return

        # 1) 메인 스텝 행 찾기(위로 올라가며 숫자 step)
        main_row = any_row
        while main_row >= 0:
            item = self.tableWidget.item(main_row, ColNum.STEP.value)
            if item and item.text().strip().isdigit():
                break
            main_row -= 1
        if main_row < 0:
            return

        # 2) 테이블 내 모든 메인 스텝 행 목록(숫자 step이 있는 행)
        main_rows = []
        for r in range(total_rows):
            it = self.tableWidget.item(r, ColNum.STEP.value)
            if it and it.text().strip().isdigit():
                main_rows.append(r)
        if not main_rows:
            return

        # 3) 마지막 메인 스텝 여부 확인 → 마지막이면 삭제 금지
        #    (요구사항: 마지막 step이면 delete는 작동하지 않음)
        #    main_rows는 오름차순. main_row가 마지막 요소면 금지.
        if main_row == main_rows[-1]:
            # 마지막 메인 스텝 → 삭제하지 않음
            print("[INFO] 마지막 메인 스텝은 삭제하지 않습니다.")
            return

        # 4) 블록 끝(구분자 포함) 찾기
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

        # 5) 테이블에서 블록 삭제(뒤에서부터)
        for r in reversed(delete_rows):
            self.tableWidget.removeRow(r)

        # 6) 내부 리스트에서 해당 범위(메인+서브)만 제거
        self.step_list = [s for s in self.step_list if not (main_row <= s[0] <= block_end)]
        self.box_list  = [b for b in self.box_list  if not (main_row <= b[0] <= block_end)]

        # 7) 이후 행 인덱스 보정(삭제된 행 수만큼 당김)
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

        # 8) 번호/분기 재정렬
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

    def paste_step_above(self, target_row: int, copied_step: "Step"):
        """현재 row 기준 블록 위에 새 메인 스텝을 만들고 copied_step 내용을 채운다."""
        self.tableWidget.blockSignals(True)
        try:
            # 1) 구조만 먼저
            insert_pos = self.insert_step_above(target_row, do_reset=False)

            # 2) 데이터 채우기 (isLoad=1 로 자동 추가/재정렬 억제)
            # 타입
            self.new_cell(insert_pos, ColNum.TYPE.value, method_name=self.generate_combo_box)
            self.__change_combo_box(insert_pos, ColNum.TYPE.value,
                                    copied_step.type_.name, self.change_type, isLoad=1)

            # 모드
            if getattr(copied_step, "mode", None) is not None:
                self.new_cell(insert_pos, ColNum.MODE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(insert_pos, ColNum.MODE.value,
                                        copied_step.mode.name, self.change_mode, isLoad=1)

            # 조건/연산자
            if getattr(copied_step, "condition", None) is not None:
                self.new_cell(insert_pos, ColNum.END_TYPE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(insert_pos, ColNum.END_TYPE.value,
                                        copied_step.condition.name, self.change_condition, isLoad=1)
            if getattr(copied_step, "operator", None) is not None:
                self.new_cell(insert_pos, ColNum.OP.value, method_name=self.generate_combo_box)
                self.__change_combo_box(insert_pos, ColNum.OP.value,
                                        str(copied_step.operator), self.change_operator, isLoad=1)

            # 리포트 타입
            if getattr(copied_step, "report", None) is not None:
                self.new_cell(insert_pos, ColNum.REPORT_TYPE.value, method_name=self.generate_combo_box)
                self.__change_combo_box(insert_pos, ColNum.REPORT_TYPE.value,
                                        copied_step.report.name, self.change_report, isLoad=1)

            # 값들(숫자/문자 그대로)
            if getattr(copied_step, "mode_val", None) is not None:
                self.__change_value(insert_pos, ColNum.MODE_VALUE.value, copied_step.mode_val)
            if getattr(copied_step, "condition_val", None) is not None:
                self.__change_value(insert_pos, ColNum.OP_VALUE.value, copied_step.condition_val)
            if getattr(copied_step, "report_val", None) is not None:
                self.__change_value(insert_pos, ColNum.REPORT_VALUE.value, copied_step.report_val)
            if getattr(copied_step, "notetext", None) is not None:
                self.__change_value(insert_pos, ColNum.STEP_NOTE.value, copied_step.notetext)

            # Sub 복사 케이스는 여기서 만들지 않음(요구사항: 메인만). 필요하면 의존/서브도 별도 로직으로.

            # 3) 한 번만 재계산
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
        test_action = menu.addAction("tests")

        action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        row = self.tableWidget.indexAt(pos).row()

        if action == copy_action:
            self.__copy_step(row)
        elif action == paste_action:
            if self.paste_list:
                self.paste_step_above(row, self.paste_list[0][1])
        elif action == insert_action:
            self.insert_step_above(row)
        elif action == delete_action:
            self.delete_step_block(row)

    def __find_block_bounds(self, any_row: int):
        """
        any_row가 속한 '메인 스텝 블록(메인 + 서브들)'의 시작/끝 행을 반환.
        반환값: (main_row, last_step_row)
        - last_step_row는 '구분자(blank)'는 포함하지 않음.
        """
        total = self.tableWidget.rowCount()
        if total == 0:
            return None, None

        # 위로 올라가며 메인 스텝(숫자 step) 찾기
        main_row = any_row
        while main_row >= 0:
            it = self.tableWidget.item(main_row, ColNum.STEP.value)
            if it and it.text().strip().isdigit():
                break
            main_row -= 1
        if main_row < 0:
            return None, None

        # 아래로 내려가며 블록 끝 찾기 (다음 메인 스텝 직전까지)
        last_step_row = main_row
        for r in range(main_row + 1, total):
            it = self.tableWidget.item(r, ColNum.STEP.value)
            # 다음 메인 스텝(숫자) 나오면 직전이 끝
            if it and it.text().strip().isdigit():
                break
            # 공백 구분자면 step_list에는 없음 → 여기서 종료
            if not it or it.text().strip() == "":
                break
            last_step_row = r

        return main_row, last_step_row

    def __collect_block_steps(self, main_row: int, last_row: int):
        """
        step_list에서 [main_row..last_row] 구간의 Step 객체들을 깊은 복사하여 반환.
        반환: [(row, Step_copy), ...]  (row는 원본 row, 참고용)
        """
        result = []
        rows = set(range(main_row, last_row + 1))
        for row, step in self.step_list:
            if row in rows and hasattr(step, "type_"):
                result.append((row, copy.deepcopy(step)))
        return result


    def __bulk_insert_rows(self, insert_pos: int, count: int, add_separator: bool = True):
        """
        insert_pos 위치에 count개의 '스텝 행'을 한꺼번에 삽입한다.
        - step_list / box_list의 row 인덱스를 모두 +count 보정
        - tableWidget에도 실제 행 삽입
        - 각 셀은 기본 구조만 만들어둔다(콤보/값 셀). 값 채우기는 이후 수행.

        add_separator: True면 블록 뒤에 구분자(blank) 1행을 추가.
        """
        # 1) 내부 인덱스 보정
        for s in self.step_list:
            if s[0] >= insert_pos:
                s[0] += count
        for b in self.box_list:
            if b[0] >= insert_pos:
                b[0] += count
        for s in self.step_list:
            st = s[1]
            if hasattr(st, "depend_row") and st.depend_row is not None and st.depend_row >= insert_pos:
                st.depend_row += count

        # 2) GUI 행 삽입 + step_list/box_list 엔트리 생성
        for k in range(count):
            r = insert_pos + k
            self.tableWidget.insertRow(r)
            self.new_step(r)  # step_list/box_list에 [r, Step()] / [r, Box_method_item()] 추가

            # 기본 셀 구성 (콤보/값)
            for col in (ColNum.TYPE.value, ColNum.MODE.value, ColNum.END_TYPE.value,
                        ColNum.OP.value, ColNum.REPORT_TYPE.value):
                self.new_cell(r, col, method_name=self.generate_combo_box)
            for col in (ColNum.MODE_VALUE.value, ColNum.OP_VALUE.value,
                        ColNum.GOTO.value, ColNum.REPORT_VALUE.value, ColNum.STEP_NOTE.value):
                self.new_cell(r, col, method_name=self.generate_value_cell)

        # 3) 블록 뒤 구분자 1행 (table 전용)
        if add_separator:
            blank_row = insert_pos + count
            self.tableWidget.insertRow(blank_row)
            self.new_cell(blank_row, ColNum.STEP.value, text="")



    def __apply_step_to_row(self, row: int, step_obj, main_base_row: int, isLoad=True):
        """
        딥카피된 step_obj 내용을 row에 반영.
        - Sub이면 depend_row를 main_base_row로 매핑
        - 콤보/값 셀은 load 방식으로 반영 (isLoad=True)
        """
        # 1) TYPE
        type_name = getattr(step_obj.type_, "name", None)
        if type_name:
            self.new_cell(row, ColNum.TYPE.value, text=type_name, method_name=self.generate_combo_box)
            self.__change_combo_box(row, ColNum.TYPE.value, type_name, self.find_method_box(ColNum.TYPE.value), isLoad=1)

        # 2) Sub의 depend_row 보정
        if getattr(step_obj, "type_", None) == procedure.Type.Sub:
            step_obj.depend_row = main_base_row

        # 3) MODE
        if getattr(step_obj, "mode", None):
            m = str(getattr(step_obj.mode, "name", step_obj.mode))
            self.new_cell(row, ColNum.MODE.value, text=m, method_name=self.generate_combo_box)
            self.__change_combo_box(row, ColNum.MODE.value, m, self.find_method_box(ColNum.MODE.value), isLoad=1)

        # 4) END_TYPE(Condition)
        if getattr(step_obj, "condition", None):
            c = str(getattr(step_obj.condition, "name", step_obj.condition))
            self.new_cell(row, ColNum.END_TYPE.value, text=c, method_name=self.generate_combo_box)
            self.__change_combo_box(row, ColNum.END_TYPE.value, c, self.find_method_box(ColNum.END_TYPE.value), isLoad=1)

        # 5) OP
        if getattr(step_obj, "operator", None):
            op = str(step_obj.operator)
            self.new_cell(row, ColNum.OP.value, text=op, method_name=self.generate_combo_box)
            self.__change_combo_box(row, ColNum.OP.value, op, self.find_method_box(ColNum.OP.value), isLoad=1)

        # 6) REPORT_TYPE
        if getattr(step_obj, "report", None):
            r = str(getattr(step_obj.report, "name", step_obj.report))
            self.new_cell(row, ColNum.REPORT_TYPE.value, text=r, method_name=self.generate_combo_box)
            self.__change_combo_box(row, ColNum.REPORT_TYPE.value, r, self.find_method_box(ColNum.REPORT_TYPE.value), isLoad=1)

        # 7) 값 칸들
        if getattr(step_obj, "mode_val", None) is not None:
            self.__change_value(row, ColNum.MODE_VALUE.value, step_obj.mode_val)
        if getattr(step_obj, "condition_val", None) is not None:
            self.__change_value(row, ColNum.OP_VALUE.value, step_obj.condition_val)
        if getattr(step_obj, "goto", None) is not None:
            self.__change_value(row, ColNum.GOTO.value, step_obj.goto)
        if getattr(step_obj, "report_val", None) is not None:
            self.__change_value(row, ColNum.REPORT_VALUE.value, step_obj.report_val)
        if getattr(step_obj, "notetext", None) is not None:
            self.__change_value(row, ColNum.STEP_NOTE.value, step_obj.notetext)


    def _cell_text(self, row, col):
        item = self.tableWidget.item(row, col)
        return item.text().strip() if item else ""

    def _is_int_text(self, s: str) -> bool:
        return s.isdigit()

    def _safe_step_number(self, row):
        txt = self._cell_text(row, ColNum.STEP.value)
        return int(txt) if self._is_int_text(txt) else None

    def _set_step_text(self, row, text):
        # step 번호 셀에 직접 텍스트 씁니다.
        self.new_cell(row, ColNum.STEP.value, text=str(text) if text is not None else "")
