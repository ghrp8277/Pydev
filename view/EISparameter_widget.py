from __future__ import annotations
from typing import Dict, Any
from PySide6.QtCore import Qt, QLocale
from PySide6.QtGui import QDoubleValidator, QIntValidator, QIcon
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QWidget, QFileDialog, QMessageBox
from view.ui_EISset_Widget import Ui_Form
from model import file
import os

def _to_float(text: str) -> float:
    if text is None:
        return 0.0
    return float(text.replace(',', '.').strip())

def _to_int(text: str) -> int:
    if text is None or text.strip() == '':
        return 0
    return int(text.strip())


class EISSetWidget(QWidget, Ui_Form):
    eis_saved = Signal(dict)  # EIS 설정 데이터 송신용 시그널
    def __init__(self, StepNum=None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setupUi(self)
        self.step_num = StepNum
        self.file_name = None
        self.load_step = None

        # ---------- 🔹 Validator 설정 ----------
        # 소수점 3자리까지 허용 (0 이상)
        dbl3 = QDoubleValidator(self)
        dbl3.setBottom(0.0)
        dbl3.setDecimals(3)
        dbl3.setNotation(QDoubleValidator.StandardNotation)

        # 정수만 허용
        int_validator = QIntValidator(self)
        int_validator.setBottom(0)

        # 적용
        self.lineEdit.setValidator(dbl3)
        self.lineEdit_2.setValidator(dbl3)
        self.lineEdit_3.setValidator(dbl3)
        self.lineEdit_4.setValidator(int_validator)

        self.lineEdit.editingFinished.connect(lambda: self._format_decimal(self.lineEdit))
        self.lineEdit_2.editingFinished.connect(lambda: self._format_decimal(self.lineEdit_2))
        self.lineEdit_3.editingFinished.connect(lambda: self._format_decimal(self.lineEdit_3 , 6))
        self.lineEdit_4.editingFinished.connect(lambda: self._format_integer(self.lineEdit_4))
        # 로케일 '.' 강제
        QLocale.setDefault(QLocale.c())

        # ---------- 🔹 아이콘 설정 ----------
        base_dir = os.path.dirname(__file__)
        icon_dir = os.path.join(base_dir, "icons")

        self.pushButton.setIcon(QIcon(os.path.join(icon_dir, "open.png")))
        self.pushButton_2.setIcon(QIcon(os.path.join(icon_dir, "saveColor.png")))

        # ---------- 🔹 시그널 연결 ----------
        self.pushButton.clicked.connect(self._on_load)
        self.pushButton_2.clicked.connect(self._on_save)
        self.comboBox.currentTextChanged.connect(self._on_mode_changed)

        # 초기 단위 설정
        self._on_mode_changed(self.comboBox.currentText())

    # ======================================================
    # 데이터 입출력
    # ======================================================
    def _format_decimal(self, line_edit, decnum = 3):
        """입력 후 자동으로 소수점 3자리 포맷 적용"""
        text = line_edit.text().strip()
        if not text:
            return
        try:
            value = float(text)
            line_edit.setText(f"{value:.{decnum}f}")
        except ValueError:
            pass

    def _format_integer(self, line_edit):
        """정수 입력 자동 포맷"""
        text = line_edit.text().strip()
        if not text:
            return
        try:
            value = int(float(text))
            line_edit.setText(str(value))
        except ValueError:
            pass

    def get_params(self) -> Dict[str, Any]:
        return {
            "mode": self.comboBox.currentText(),
            "start_frequency_hz": _to_float(self.lineEdit.text()),
            "stop_frequency_hz":  _to_float(self.lineEdit_2.text()),
            "amplitude":          _to_float(self.lineEdit_3.text()),
            "point_number":       _to_int(self.lineEdit_4.text()),
        }

    def set_params(self, params: Dict[str, Any]) -> None:
        mode = str(params.get("mode", "GEIS"))
        idx = self.comboBox.findText(mode)
        self.comboBox.setCurrentIndex(idx if idx >= 0 else 0)

        # 문자열로 변환 후 표시 (setText(int) 오류 방지)
        self.lineEdit.setText(f"{float(params.get('start_frequency_hz', 100000.000)):.3f}")
        self.lineEdit_2.setText(f"{float(params.get('stop_frequency_hz', 0.100)):.3f}")
        self.lineEdit_3.setText(f"{float(params.get('amplitude', 0.010000)):.3f}")
        self.lineEdit_4.setText(str(int(params.get('point_number', 50))))

    # ======================================================
    # 파일 저장 / 불러오기
    # ======================================================
    def _on_save(self) -> None:
        try:
            # ① 저장 경로 선택
            path, _ = QFileDialog.getSaveFileName(
                self, "EIS 설정 저장", "", "Excel 파일 (*.xlsx)"
            )
            if not path:
                return
            eis_data = [[
                #self.step_num,
                self.comboBox.currentText(),
                self.lineEdit.text(),
                self.lineEdit_2.text(),
                self.lineEdit_3.text(),
                self.lineEdit_4.text()
            ]]
            file.save_eis(path, eis_data)
            self.file_name = path

            eis_data = [[
                #self.step_num,
                self.comboBox.currentText(),
                self.lineEdit.text(),
                self.lineEdit_2.text(),
                self.lineEdit_3.text(),
                self.lineEdit_4.text(),
                path
            ]]
            self.eis_saved.emit(eis_data)
            QMessageBox.information(self, "저장 완료", f"EIS 설정이 저장되었습니다:\n{path}")

        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"파일 저장 중 오류가 발생했습니다.\n\n{e}")

    def _on_load(self) -> None:
        try:
            path, _ = QFileDialog.getOpenFileName(
                self, "EIS 설정 불러오기", "", "Excel 파일 (*.xlsx)"
            )
            if not path:
                return

            # ① Excel 파일 읽기
            eis_data = file.open_file(path)  # 2D 리스트 반환
            if not eis_data or len(eis_data[0]) < 5:
                raise ValueError("EIS 설정 파일 형식이 올바르지 않습니다.")

            # ② 첫 번째 행을 기준으로 UI에 반영
            self.comboBox.setCurrentText(str(eis_data[0][0]))
            self.lineEdit.setText(str(eis_data[0][1]))
            self.lineEdit_2.setText(str(eis_data[0][2]))
            self.lineEdit_3.setText(str(eis_data[0][3]))
            self.lineEdit_4.setText(str(eis_data[0][4]))

            self.file_name = path
            eis_data[0].append(path)  # 경로 추가
            self.eis_saved.emit(eis_data)
            QMessageBox.information(self, "불러오기 완료", f"EIS 설정을 불러왔습니다:\n{path}")

        except Exception as e:
            QMessageBox.critical(self, "불러오기 실패", f"파일 불러오기 중 오류가 발생했습니다.\n\n{e}")

    # ======================================================
    # GEIS / PEIS 모드에 따른 단위 변경
    # ======================================================
    def _on_mode_changed(self, mode_text: str) -> None:
        """
        GEIS 선택 시 label_8 → 'A'
        PEIS 선택 시 label_8 → 'V'
        """
        if mode_text.upper() == "GEIS":
            self.label_8.setText("A")
        elif mode_text.upper() == "PEIS":
            self.label_8.setText("V")
        elif mode_text.upper() == "ACIR":
            self.label_8.setText("A")
            self.lineEdit_2.setText("0")
            self._format_decimal(self.lineEdit_2 , 3)
            self.lineEdit_2.setDisabled(True)
        else:
            self.label_8.setText("?")
