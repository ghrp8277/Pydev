
import os
from PySide6.QtWidgets import QMainWindow
from PySide6.QtGui import QIcon, QKeySequence
from PySide6.QtCore import Qt
from view.ui.main_ui import MainUI

class MainView(QMainWindow, MainUI):
    def __init__(self, controller=None):
        super().__init__()
        self.setupUi(self)
        self.controller = controller

        # 아이콘과 단축키 적용
        self._apply_icons()
        self._apply_shortcuts()

    def _icon_path(self, name: str) -> str:
        src_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        return os.path.join(src_dir,"src", "public", "icons", name)
    
    def _apply_icons(self):
        mapping = {
            'actionOpen': 'open.png',
            'actionSave': 'saveColor.png',
            'actionNew': 'new.png',
            'actionCut': 'cut.png',
            'actionCopy': 'copy.png',
            'actionPaste': 'paste.png',
            'actionViewMode': 'windowMode.png',
            'actionCascade': 'cascade.png',
            'actionTile': 'tile.png',
            'actionSave_all': 'saveAllColor.png',
            'actionSetting': 'setting.png'
        }
        for attr, icon_name in mapping.items():
            act = getattr(self, attr, None)
            if act is not None:
                path = self._icon_path(icon_name)
                act.setIcon(QIcon(path))

    def _apply_shortcuts(self):
        if getattr(self, 'actionNew', None):
            self.actionNew.setShortcut(QKeySequence('Ctrl+N'))
        if getattr(self, 'actionOpen', None):
            self.actionOpen.setShortcut(QKeySequence('Ctrl+O'))
        if getattr(self, 'actionSave', None):
            self.actionSave.setShortcut(QKeySequence('Ctrl+S'))
        if getattr(self, 'actionCopy', None):
            self.actionCopy.setShortcut(QKeySequence('Ctrl+C'))
        if getattr(self, 'actionPaste', None):
            self.actionPaste.setShortcut(QKeySequence('Ctrl+V'))
        if getattr(self, 'actionCut', None):
            self.actionCut.setShortcut(QKeySequence('Ctrl+X'))
