from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import QFileDialog, QMdiSubWindow

from view.ui.main_ui import MainUI

from controller.table_controller import TableController
from view.table_view import TableView

class MainController:
    def __init__(self, view: MainUI):
        self.view = view
        self.view.controller = self

        self._new_procedure_count = 1
        self._windows: list[QMdiSubWindow] = []

        self._configure_actions()
        #self._initialize_default_window()

    def _configure_actions(self) -> None:
        toolbar = self.view
        toolbar.actionNew.triggered.connect(self.new_table)
        toolbar.actionOpen.triggered.connect(self.open_table)

    def new_table(self) -> None:
        title = self._next_new_title()
        self._create_table_window(title=title)

    def open_table(self) -> None:
        path, _ = QFileDialog.getOpenFileName( self.view, "Open Procedure", "", "Excel (*.xlsx)")
        if not path:
            return
        title = Path(path).name
        self._create_table_window(title=title, source_path=path)

    def _next_new_title(self) -> str:
        title = f"New Procedure {self._new_procedure_count}"
        self._new_procedure_count += 1
        return title

    def _initialize_default_window(self) -> None:
        self._create_table_window(title=self._next_new_title())

    def _create_table_window(self, *, title: str, source_path: Optional[str] = None):
        table_controller = TableController()
        table_view = TableView()
        table_view.set_controller(table_controller)
        table_controller.set_view(table_view)

        #MDI SubWindow 생성
        sub = QMdiSubWindow()
        sub.setWidget(table_view)
        sub.setWindowTitle(title)

        #Optional: 파일 경로 전달
        if source_path:
            sub.setProperty("source_path", source_path)

        self._windows.append(sub)
        self.view.mdiArea.addSubWindow(sub)
        table_controller.set_seed_data()
        sub.show()

        return sub
