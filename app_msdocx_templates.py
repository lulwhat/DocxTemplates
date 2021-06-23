import sys
import os.path
import csv
import io

from PyQt5.QtWidgets import (QApplication, QDialog,
                             QHBoxLayout, QVBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, 
                             QTableWidget, QVBoxLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
                             qApp, QSpacerItem, QHeaderView
                            )
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtCore import Qt, QEvent

from docx.opc.exceptions import PackageNotFoundError

from msdocx_templates_functionality import *

class MsDocxTemplatesGui(QDialog):
    def __init__(self, parent=None):
        super(MsDocxTemplatesGui, self).__init__(parent)

        self.template_path = os.path.expanduser("~/Documents")
        self.save_path = os.path.expanduser("~/Documents")
        spacer = QSpacerItem(0,10)
    
        self.originalPalette = QApplication.palette()

        # choose template layout
        templatePathLabel = QLabel("Выберите шаблон:")
        self.templatePathLineEdit = QLineEdit()
        templateBrowseButton = QPushButton("Обзор")
        templateBrowseButton.clicked.connect(self.templateBrowseButtonAction)

        templatePathLayout = QHBoxLayout()
        templatePathLayout.addWidget(templatePathLabel)
        templatePathLayout.addWidget(self.templatePathLineEdit)
        templatePathLayout.addWidget(templateBrowseButton)

        # save docx path layout
        savePathLabel = QLabel("Сохранить файл как:")
        self.savePathLineEdit = QLineEdit()
        savePathBrowseButton = QPushButton("Обзор")
        savePathBrowseButton.clicked.connect(self.savePathBrowseButtonAction)

        savePathLayout = QHBoxLayout()
        savePathLayout.addWidget(savePathLabel) 
        savePathLayout.addWidget(self.savePathLineEdit)
        savePathLayout.addWidget(savePathBrowseButton)

        # export layout
        self.progressBar = QProgressBar()
        self.progressBar.setValue(0)
        exportButton = QPushButton("Экспорт")
        exportButton.setStyleSheet("background-color: rgb(200, 230, 200)")
        exportButton.clicked.connect(self.exportButtonAction)

        exportLayout = QHBoxLayout()
        exportLayout.addWidget(self.progressBar)
        exportLayout.addWidget(exportButton)

        # table layout
        self.tableWidget = QTableWidget(1,2)
        self.fillTableWidgetNoneCells()
        self.tableWidget.installEventFilter(self)
        self.tableWidget.setHorizontalHeaderLabels(["Переменные", "Значения"])
        self.tableWidget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)

        self.flags = Qt.ItemFlags()
        self.flags != Qt.ItemIsSelectable
        self.flags != Qt.ItemIsEnabled
        self.tableWidget.item(0,0).setFlags(self.flags)

        tableLayout = QHBoxLayout()
        tableLayout.addWidget(self.tableWidget)


        # create main layout
        mainLayout = QVBoxLayout()
        mainLayout.addLayout(templatePathLayout)
        mainLayout.addLayout(savePathLayout)
        mainLayout.addLayout(exportLayout)
        mainLayout.addSpacerItem(spacer)
        mainLayout.addLayout(tableLayout)

        # app settings
        self.setLayout(mainLayout)
        self.setWindowTitle("Шаблоны MS Word")
        self.setMinimumSize(700, 400)
        self.setWindowIcon(QIcon(self.resource_path("logo_ug.png")))
        self.setWindowFlags(Qt.WindowFlags())

    def templateBrowseButtonAction(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(
            self,
            "Загрузить шаблон MS Word",
            self.template_path,
            "Документы MS Word (*.docx;*.doc);;Все файлы (*)",
            options=options
        )
        if fileName:
            self.template_path = os.path.dirname(fileName)
            # set file text field and change forward slash to windows one
            self.templatePathLineEdit.setText(fileName.replace("/", "\\"))

            # read template and fill table with variables
            try:
                self.docxHandler = DocxHandler(fileName)
                self.replacement_data = self.docxHandler.templateRead()
                if len(self.replacement_data) > 0:
                    self.tableWidget.setRowCount(0)
                    vars = list(self.replacement_data.keys())
                    for i in range(len(vars)):
                        self.tableWidget.insertRow(self.tableWidget.rowCount())
                        self.tableWidget.setItem(i,0, QTableWidgetItem(vars[i]))
                        self.tableWidget.item(i,0).setFlags(self.flags)
                    self.fillTableWidgetNoneCells()
                else:
                    msg = QMessageBox()
                    msg.warning(self, "Ошибка", "Переменные в файле не найдены")
            except PackageNotFoundError:
                msg = QMessageBox()
                msg.warning(self, "Ошибка", "Неверный формат шаблона")

    def savePathBrowseButtonAction(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить файл как",
            self.save_path + "/*.docx",
            "Документы MS Word (*.docx;*.doc);;Все файлы (*)",
            options=options
        )
        if fileName:
            self.save_path = os.path.dirname(fileName)
            # set file text field and change forward slash to windows one
            self.savePathLineEdit.setText(fileName.replace("/", "\\"))

    # the export
    def exportButtonAction(self):
        self.progressBar.setValue(0)
        try:
            for row in range(self.tableWidget.rowCount()):
                self.replacement_data[self.tableWidget.item(row,0).text()] = self.tableWidget.item(row,1).text()
            self.docxHandler.docxReplace(self.templatePathLineEdit.text(), self.replacement_data)
            if self.docxHandler.docxSave(self.savePathLineEdit.text()) == True:
                QMessageBox.about(self, "Шаблоны Ms Word", f"Файл {os.path.basename(self.savePathLineEdit.text())} создан")
                self.progressBar.setValue(100)
        except FileNotFoundError:
            msg = QMessageBox()
            msg.warning(self, "Ошибка", "Введите имя файла для сохранения")
        except AttributeError:
            msg = QMessageBox()
            msg.warning(self, "Ошибка", "Не выбран файл шаблона")
        except SavePathIsNotAbsoluteError:
            msg = QMessageBox()
            msg.warning(self, "Ошибка", "Введите полный путь для сохранения файла")
        except SaveFileWrongFormatError:
            msg = QMessageBox()
            msg.warning(self, "Ошибка", "Неверный формат экспортируемого файла")
        except PackageNotFoundError:
            msg = QMessageBox()
            msg.warning(self, "Ошибка", "Неверный формат шаблона")



    def fillTableWidgetNoneCells(self):
        for row in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                if self.tableWidget.item(row, col) == None:
                    self.tableWidget.setItem(row, col, QTableWidgetItem(""))

    # event filters for table widget
    def eventFilter(self, source, event):
        if (event.type() == QEvent.KeyPress and
            event.matches(QKeySequence.Copy)):
            self.copySelection()
            return True
        if (event.type() == QEvent.KeyPress and
            event.matches(QKeySequence.Paste)):
            self.pasteSelection()
            return True
        if (event.type() == QEvent.KeyPress and
            event.matches(QKeySequence.Delete or
            event.matches(QKeySequence.Backspace))):
            self.deleteSelection()
            return True
        if (event.type() == QEvent.KeyPress and
            event.matches(QKeySequence.Cut)):
            self.cutSelection()
            return True
        return super(MsDocxTemplatesGui, self).eventFilter(source, event)

    # CTRL+C table widget event
    def copySelection(self):
        selection = self.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            csv.writer(stream, delimiter='\t').writerows(table)
            qApp.clipboard().setText(stream.getvalue())
        return

    # CTRL+V table widget event
    def pasteSelection(self):
        selection = self.tableWidget.selectedIndexes()
        if selection:
            model = self.tableWidget.model()

            buffer = qApp.clipboard().text()
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            reader = csv.reader(io.StringIO(buffer), delimiter='\t')
            if len(rows) == 1 and len(columns) == 1:
                for i, line in enumerate(reader):
                    for j, cell in enumerate(line):
                        model.setData(model.index(rows[0]+i,columns[0]+j), cell)
            else:
                arr = [ [ cell for cell in row ] for row in reader]
                for index in selection:
                    row = index.row() - rows[0]
                    column = index.column() - columns[0]
                    try:
                        model.setData(model.index(index.row(), index.column()), arr[row][column])
                    except IndexError:
                        pass
        return

    # CTRL+X table widget event
    def cutSelection(self):
        selection = self.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            csv.writer(stream, delimiter='\t').writerows(table)
            qApp.clipboard().setText(stream.getvalue())
            for row in rows:
                for col in columns:
                    self.tableWidget.setItem(row, col, QTableWidgetItem(""))
        return

    # delete/backspace table widget event
    def deleteSelection(self):
        selection = self.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            for row in rows:
                for col in columns:
                    self.tableWidget.setItem(row, col, QTableWidgetItem(""))

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    gallery = MsDocxTemplatesGui()
    gallery.show()
    sys.exit(app.exec_())