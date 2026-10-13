"""python -m efield2d.app [file.dxf | project.efp]"""
import sys

from PySide6.QtWidgets import QApplication

from .mainwindow import MainWindow


def main(argv=None):
    argv = list(sys.argv if argv is None else argv)
    app = QApplication(argv)
    win = MainWindow()
    win.show()
    if len(argv) > 1:
        path = argv[1]
        (win.open_project if path.lower().endswith((".efp", ".json")) else win.open_dxf)(path)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
