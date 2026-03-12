from PyQt6.QtWidgets import QTableWidget, QStyledItemDelegate, QLineEdit

class CleanEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setStyleSheet("""
            QLineEdit {
                background-color: white;
                color: black;
            }
        """)
        return editor

class DragDropTableWidget(QTableWidget):
    def __init__(self, *args, load_callback=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDragDropMode(QTableWidget.DragDropMode.DropOnly)
        self.load_callback = load_callback

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if self.load_callback:
                    self.load_callback(file_path)
            event.acceptProposedAction()
        else:
            event.ignore()