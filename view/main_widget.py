from PySide6.QtWidgets import (QMainWindow, QMdiArea, QFileDialog,
    QLabel)

from view import ui_BuildTestMainWindow
from view import table_widget
from view import EISparameter_widget
from PySide6.QtCore import Qt

# !void func for execution, raise error
def not_yet() -> None: ...
    # raise Exception("Change this function to the actual!")
# !====================================

#====================Main Window==========================#
class MainWindow(QMainWindow, ui_BuildTestMainWindow.Ui_BuidTestMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setupUi(self)

        self.eis_window = None

        # save opened file name
        self.open_file_list = []
        self.status_step = QLabel()
        self.status_step.setMaximumSize(100, 100)
        self.status_step.setText("Step: 1")
        self.status_row = QLabel()
        self.status_row.setMaximumSize(100, 100)
        self.status_row.setText("Row: 0")
        self.status_column = QLabel()
        self.status_column.setMaximumSize(100, 100)
        self.status_column.setText("Column: 0")

        self.statusBar.addWidget(self.status_step, 1)
        self.statusBar.addWidget(self.status_row, 1)
        self.statusBar.addWidget(self.status_column, 1)

        #----- Signal & Setting -----#

        #----- Menu & Tool bar action -----#
        # format: actionName.triggered.connect(funcName)

        #----- File Menu -----#
        # TODO: connect with slot -> open, save, save_all, save_as
        self.actionNew.triggered.connect(lambda: self.new_sub_window())
        self.actionOpen.triggered.connect(self.open_file_dialog)
        self.actionSave.triggered.connect(self.save_file_dialog)
        self.actionSave_all.triggered.connect(not_yet)
        self.actionSave_AS.triggered.connect(not_yet)

        #----- Edit Menu -----#
        self.actionUndo.triggered.connect(not_yet)
        self.actionCut.triggered.connect(not_yet)
        self.actionCopy.triggered.connect(not_yet)
        self.actionPaste.triggered.connect(not_yet)
        self.actionInsertStep.triggered.connect(not_yet)
        self.actionDeleteStep.triggered.connect(not_yet)
        self.actionClearCell.triggered.connect(not_yet)

        #----- Option Menu -----#
        self.actionSetting.triggered.connect(not_yet)

        #----- Window Menu -----#
        self.actionViewMode.triggered.connect(self.change_view_mode)
        self.actionCascade.triggered.connect(self.align_cascade)
        self.actionTile.triggered.connect(self.align_tile)

        #----- Help Menu -----#
        self.actionAbout.triggered.connect(not_yet)


    def open_eis_window(self, step_num=None):
        if self.eis_window is not None and self.eis_window.isVisible():
            self.eis_window.activateWindow()
            return

        self.eis_window = EISparameter_widget.EISSetWidget(StepNum=step_num, parent=self)
        self.eis_window.setWindowTitle("EIS 설정")
        self.eis_window.setWindowFlags(Qt.Window)
        self.eis_window.setFixedSize(180, 300)
        self.eis_window.setAttribute(Qt.WA_DeleteOnClose)
        self.eis_window.destroyed.connect(lambda: setattr(self, "eis_window", None))
        self.eis_window.eis_saved.connect(self.handle_eis_saved)
        self.eis_window.show()
    #===== Slots =====#
    def handle_eis_saved(self, eis_data: dict):
        """MainWindow에서 TableWidget으로 데이터 전달"""
        active_window = self.mdiArea.currentSubWindow()
        if not active_window:
            return
        table = active_window.widget()
        if hasattr(table, "add_eis_parameters"):
            table.add_eis_parameters(eis_data)
        print("EIS 설정이 TableWidget으로 전달됨:", eis_data)
        
    #----- File Menu method-----#
    def open_file_dialog(self):
        file_name = (QFileDialog.getOpenFileName(self, "Open File", '', 'xlsx(*.xlsx)'))[0]
        if file_name:
            self.new_sub_window(file_name)
            self.open_file_list.append(file_name)   # file관리를 위해 list 설정
            
    def handleActivationChange(self, subwindow):
        if subwindow is self.parent():
            print ('activated:', self)
        else:
            print ('deactivated:', self)
            
    def save_file_dialog(self, index=-1):
        self.mdiArea.currentSubWindow().close()
        #print(self.mdiArea.currentSubWindow().__init__.__dir__)
        """
        if index == -1:
            file_name = (QFileDialog.getSaveFileName(self, "Save File", '', 'xlsx(*.xlsx)'))[0]
        else:
            file_name = self.open_file_list[index]
        """
    #----- mdiArea -----#

    # generate new sub window (table widget)
    def new_sub_window(self, file_name=None):
        
        sub_window = table_widget.TableWidget(
            (self.status_step, self.status_row, self.status_column), file_name)
        #print(sub_window)
        sub_window_count = len(self.mdiArea.subWindowList())
        if file_name:
            sub_window.setWindowTitle(file_name)
        elif sub_window_count != 0:
            sub_window.setWindowTitle(f"NewProcedure({sub_window_count})")
        self.mdiArea.addSubWindow(sub_window)
        sub_window.show()

    # change view mode -> window | tab
    def change_view_mode(self):
        tab_mode = QMdiArea.ViewMode.TabbedView
        win_mode = QMdiArea.ViewMode.SubWindowView
        if self.mdiArea.viewMode()==win_mode:
            self.mdiArea.setViewMode(tab_mode)
            self.actionViewMode.setText(u"TabMode")
            self.mdiArea.setTabsClosable(True)
            self.mdiArea.setTabsMovable(True)
        else:
            self.mdiArea.setViewMode(win_mode)
            self.actionViewMode.setText(u"WindowMode")

    #----- align window -----#
    def align_cascade(self):
        self.mdiArea.cascadeSubWindows()

    def align_tile(self):
        self.mdiArea.tileSubWindows()

    def closeEvent(self, event):
        if hasattr(self, "eis_window") and self.eis_window is not None:
            self.eis_window.close()
        super().closeEvent(event)