import sys
import os
from PySide6.QtWidgets import QApplication


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_PATH = os.path.join(BASE_DIR, "src")
if SRC_PATH not in sys.path:
    sys.path.insert(0, SRC_PATH)

print("[DEBUG] sys.path:", sys.path[:3])

try:
    from view.main_view import MainView
    from controller.main_controller import MainController
except ModuleNotFoundError as e:
    print(f"[ERROR] Import 실패: {e}")
    sys.exit(1)

def main():
    app = QApplication(sys.argv)
    window = MainView()
    MainController(window)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()