
from PySide6.QtWidgets import QApplication

from view import main_widget

#===== app 실행 부분 =====#
if __name__ == '__main__':
    app = QApplication()
    window = main_widget.MainWindow()
    window.show()

    app.exec()