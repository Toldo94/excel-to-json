import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QMessageBox, QPlainTextEdit, QPushButton, QSpinBox,
    QVBoxLayout, QWidget,
)

from .converter import convert_workbook, sheet_names, to_json_text, write_json


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to JSON")
        self.resize(900, 600)
        self.setAcceptDrops(True)
        self.path = None
        self.data = None

        self.drop_label = QLabel("Drop an .xlsx file here, or click Open…")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet("border: 2px dashed gray; padding: 14px; border-radius: 8px;")
        open_btn = QPushButton("Open…")
        open_btn.clicked.connect(self.pick_file)

        self.sheets = QListWidget()
        self.sheets.setMaximumWidth(220)
        self.sheets.itemChanged.connect(self.refresh)
        self.header_row = QSpinBox()
        self.header_row.setMinimum(1)
        self.header_row.setMaximum(1000)
        self.header_row.valueChanged.connect(self.refresh)
        self.drop_nulls = QCheckBox("Omit empty fields")
        self.drop_nulls.toggled.connect(self.refresh)

        side = QVBoxLayout()
        side.addWidget(QLabel("Sheets"))
        side.addWidget(self.sheets)
        side.addWidget(QLabel("Header row"))
        side.addWidget(self.header_row)
        side.addWidget(self.drop_nulls)

        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setFont(QFont("Menlo", 11))
        self.status = QLabel("")

        self.save_btn = QPushButton("Save JSON…")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save)
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(self.preview.toPlainText()))

        top = QHBoxLayout()
        top.addWidget(self.drop_label, 1)
        top.addWidget(open_btn)
        mid = QHBoxLayout()
        mid.addLayout(side)
        mid.addWidget(self.preview, 1)
        bottom = QHBoxLayout()
        bottom.addWidget(self.status, 1)
        bottom.addWidget(self.copy_btn)
        bottom.addWidget(self.save_btn)

        root = QVBoxLayout(self)
        root.addLayout(top)
        root.addLayout(mid, 1)
        root.addLayout(bottom)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            self.load(urls[0].toLocalFile())

    def pick_file(self):
        f, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel (*.xlsx *.xlsm)")
        if f:
            self.load(f)

    def load(self, path):
        try:
            names = sheet_names(path)
        except Exception as ex:
            QMessageBox.critical(self, "Cannot open file", f"{ex}\n\nOnly .xlsx / .xlsm are supported.")
            return
        self.path = path
        self.drop_label.setText(path)
        self.sheets.blockSignals(True)
        self.sheets.clear()
        for n in names:
            item = QListWidgetItem(n)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.sheets.addItem(item)
        self.sheets.blockSignals(False)
        self.refresh()

    def selected(self):
        return [self.sheets.item(i).text() for i in range(self.sheets.count())
                if self.sheets.item(i).checkState() == Qt.Checked]

    def refresh(self, *_):
        if not self.path:
            return
        try:
            self.data = convert_workbook(self.path, self.selected(), self.header_row.value(),
                                         drop_null_fields=self.drop_nulls.isChecked())
        except Exception as ex:
            self.preview.setPlainText(f"Error: {ex}")
            self.save_btn.setEnabled(False)
            self.copy_btn.setEnabled(False)
            return
        text = to_json_text(self.data)
        self.preview.setPlainText(text)
        total = sum(len(v) for v in self.data.values())
        self.status.setText(f"{len(self.data)} sheet(s), {total} row(s)")
        ok = bool(self.data)
        self.save_btn.setEnabled(ok)
        self.copy_btn.setEnabled(ok)

    def save(self):
        default = str(Path(self.path).with_suffix(".json"))
        f, _ = QFileDialog.getSaveFileName(self, "Save JSON", default, "JSON (*.json)")
        if f:
            write_json(self.data, f)
            self.status.setText(f"Saved {f}")


def main():
    app = QApplication(sys.argv)
    w = Window()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
