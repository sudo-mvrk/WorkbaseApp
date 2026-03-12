from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidgetItem,
    QPushButton, QLabel, QLineEdit, QFileDialog, QMessageBox, QHeaderView, QToolBar
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeySequence, QAction

from workbase.utils import parse_number
from workbase.ui_components import DragDropTableWidget, CleanEditDelegate
from workbase.io_handlers import save_json, load_file, export_txt, export_excel


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Layouts")
        self.current_file = None
        self.resize(1100, 550)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout()
        central.setLayout(layout)

        self._setup_toolbar()

        self.table = DragDropTableWidget(0, 4, load_callback=self.load_file_path)
        self.table.setHorizontalHeaderLabels(
            ["Название макета", "Цена за штуку", "Кол-во экземпляров", "Общая стоимость"])

        header = self.table.horizontalHeader()
        self.table.setColumnWidth(0, 630)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        for i in range(1, 3):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)

        self.table.setItemDelegate(CleanEditDelegate())
        layout.addWidget(self.table)

        self._setup_bottom_panel(layout)
        self.table.cellChanged.connect(self.on_cell_changed)
        self.add_row()

    def _setup_toolbar(self):
        save_over_act = QAction("Сохранить", self)
        save_over_act.setShortcut(QKeySequence("Ctrl+S"))
        save_over_act.triggered.connect(self.save_to_file)

        save_act = QAction("Сохранить как...", self)
        save_act.setShortcut(QKeySequence("Ctrl+Shift+S"))
        save_act.triggered.connect(lambda: self.save_to_file(force_dialog=True))

        load_act = QAction("Загрузить...", self)
        load_act.triggered.connect(self.load_from_file)

        export_act = QAction("Экспорт отчёта (TXT)...", self)
        export_act.triggered.connect(self.export_report)

        exel_act = QAction("Экспорт отчёта (Exel)...", self)
        exel_act.triggered.connect(self.export_excel_report)

        clear_act = QAction("Очистить таблицу", self)
        clear_act.setShortcut(QKeySequence("Ctrl+L"))
        clear_act.triggered.connect(self.clear_table)

        exit_act = QAction("Выход", self)
        exit_act.triggered.connect(self.close)

        toolbar = QToolBar("Main")
        toolbar.addAction(save_over_act)
        toolbar.addAction(save_act)
        toolbar.addAction(load_act)
        toolbar.addAction(export_act)
        toolbar.addAction(exel_act)
        toolbar.addAction(clear_act)
        toolbar.addAction(exit_act)
        self.addToolBar(toolbar)

    def _setup_bottom_panel(self, layout):
        bottom = QHBoxLayout()
        layout.addLayout(bottom)

        btn_add = QPushButton("+")
        btn_add.setToolTip("Добавить запись")
        btn_add.setFixedSize(56, 56)
        btn_add.setStyleSheet("font-size:28px; color:white; background-color:#2ecc71; border-radius:6px;")
        btn_add.clicked.connect(self.add_row)

        btn_remove = QPushButton("−")
        btn_remove.setToolTip("Удалить выбранные записи")
        btn_remove.setFixedSize(56, 56)
        btn_remove.setStyleSheet("font-size:28px; color:white; background-color:#e74c3c; border-radius:6px;")
        btn_remove.clicked.connect(self.remove_selected_rows)

        bottom.addWidget(btn_add)
        bottom.addWidget(btn_remove)
        bottom.addStretch(1)

        lbl_total = QLabel("Общая сумма:")
        lbl_total.setStyleSheet("font-weight:600; color:#555;")
        self.total_edit = QLineEdit("0.00")
        self.total_edit.setReadOnly(True)
        self.total_edit.setFixedWidth(140)
        self.total_edit.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        bottom.addWidget(lbl_total)
        bottom.addWidget(self.total_edit)

    def add_row(self, name="", price=0.0, qty=0):
        row = self.table.rowCount()
        self.table.insertRow(row)

        item_name = QTableWidgetItem(name)
        self.table.setItem(row, 0, item_name)

        item_price = QTableWidgetItem("" if price == 0 else f"{price:.2f}")
        item_price.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.table.setItem(row, 1, item_price)

        item_qty = QTableWidgetItem("" if qty == 0 else str(int(qty)))
        item_qty.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.table.setItem(row, 2, item_qty)

        item_total = QTableWidgetItem(f"{price * qty:.2f}")
        item_total.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        item_total.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
        self.table.setItem(row, 3, item_total)

        self.update_overall_sum()

    def remove_selected_rows(self):
        selected = self.table.selectedItems()
        if selected:
            rows = sorted({it.row() for it in selected}, reverse=True)
            for r in rows:
                self.table.removeRow(r)
        elif self.table.rowCount() > 0:
            self.table.removeRow(self.table.rowCount() - 1)
        self.update_overall_sum()

    def on_cell_changed(self, row, col):
        if row < 0 or col not in (1, 2):
            return

        self.table.blockSignals(True)
        price = parse_number(self.table.item(row, 1).text() if self.table.item(row, 1) else "0")
        qty = parse_number(self.table.item(row, 2).text() if self.table.item(row, 2) else "0")
        total = price * qty

        total_item = self.table.item(row, 3)
        if not total_item:
            total_item = QTableWidgetItem()
            total_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            total_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, 3, total_item)

        total_item.setText(f"{total:.2f}")
        self.table.blockSignals(False)
        self.update_overall_sum()

    def update_overall_sum(self):
        total_sum = 0.0
        for r in range(self.table.rowCount()):
            price = parse_number(self.table.item(r, 1).text() if self.table.item(r, 1) else "0")
            qty = parse_number(self.table.item(r, 2).text() if self.table.item(r, 2) else "0")
            total_sum += price * qty
        self.total_edit.setText(f"{total_sum:.2f}")

    def get_table_data(self):
        data = []
        for r in range(self.table.rowCount()):
            name = self.table.item(r, 0).text() if self.table.item(r, 0) else ""
            price = parse_number(self.table.item(r, 1).text() if self.table.item(r, 1) else "0")
            qty = int(parse_number(self.table.item(r, 2).text() if self.table.item(r, 2) else "0"))
            data.append({"name": name, "price": price, "qty": qty})
        return data

    def save_to_file(self, force_dialog=False):
        fn = self.current_file

        if not fn or fn.lower().endswith(".xml") or force_dialog:
            default_dir = str(Path(fn).parent) if fn else str(Path.home())
            default_name = f"{Path(fn).stem}.json" if fn and Path(fn).suffix.lower() == ".xml" else ""

            fn, _ = QFileDialog.getSaveFileName(
                self, "Сохранить данные (JSON)", str(Path(default_dir) / default_name), "JSON Files (*.json)"
            )
            if not fn: return
            if not fn.lower().endswith(".json"): fn += ".json"

            self.current_file = fn
            self.setWindowTitle(f"Layouts - {Path(fn).name}")

        try:
            save_json(self.current_file, self.get_table_data())
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")

    def load_file_path(self, fn):
        try:
            data = load_file(fn)
            self.table.blockSignals(True)
            self.table.setRowCount(0)

            for entry in data:
                self.add_row(entry["name"], entry["price"], entry["qty"])

            self.table.blockSignals(False)
            self.current_file = fn
            self.setWindowTitle(f"Layouts - {Path(fn).name}")
            self.update_overall_sum()
            self.table.viewport().update()

        except Exception as e:
            self.table.blockSignals(False)
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл:\n{e}")

    def load_from_file(self):
        fn, _ = QFileDialog.getOpenFileName(
            self, "Загрузить данные", str(Path.home()), "JSON Files (*.json);;XML Files (*.xml)"
        )
        if fn:
            self.load_file_path(fn)

    def export_report(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", str(Path.home()), "Text Files (*.txt)")
        if not fn: return

        try:
            total = parse_number(self.total_edit.text())
            export_txt(fn, self.get_table_data(), total)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить отчёт:\n{e}")

    def export_excel_report(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт Excel", str(Path.home()), "Excel Files (*.xlsx)")
        if not fn: return
        if not fn.endswith(".xlsx"): fn += ".xlsx"

        try:
            total = parse_number(self.total_edit.text())
            export_excel(fn, self.get_table_data(), total)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить Excel:\n{e}")

    def clear_table(self):
        reply = QMessageBox.question(
            self, "Подтверждение", "Вы действительно хотите удалить все записи?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.table.setRowCount(0)
            self.current_file = None
            self.setWindowTitle("Layouts")
            self.update_overall_sum()