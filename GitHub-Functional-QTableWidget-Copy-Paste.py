import sys
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QTableWidgetItem, QTableWidget, QHeaderView


class TableWithCopy(QTableWidget):
    """
    this class extends QTableWidget
    * supports copying multiple cell's text onto the clipboard
    * formatted specifically to work with multiple-cell paste into programs
      like google sheets, excel, or numbers
      and also copying / pasting into cells by more than 1 at a time
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            copied_cells = sorted(self.selectedIndexes())

            copy_text = ''
            max_column = copied_cells[-1].column()
            for c in copied_cells:
                copy_text += self.item(c.row(), c.column()).text()
                if c.column() == max_column:
                    copy_text += '\n'
                else:
                    copy_text += '\t'
            QApplication.clipboard().setText(copy_text)

        if event.key() == Qt.Key_C and (event.modifiers() & Qt.ControlModifier):
            self.copied_cells = sorted(self.selectedIndexes())
        elif event.key() == Qt.Key_V and (event.modifiers() & Qt.ControlModifier):
            r = self.currentRow() - self.copied_cells[0].row()
            c = self.currentColumn() - self.copied_cells[0].column()
            for cell in self.copied_cells:
                self.setItem(cell.row() + r, cell.column() + c, QTableWidgetItem(cell.data()))

# create basic UI
class Actions(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setGeometry(400, 400, 370, 560)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)

        self.setWindowTitle("TableWidget With Features")

        self.label = QLabel("TableWidget Features:\n    -  Can Copy cells and paste to things like excel\n     "
                            "- Can click on a cell and obtain cell value\n     "
                            "-  Can change the data in the cells\n     - Functionally looks good", self)
        self.label.adjustSize()
        self.label.move(50, 25)

        self.button = QPushButton(self)
        self.button.setText("Add Data 1")
        self.button.adjustSize()
        self.button.clicked.connect(self.onButtonClick)
        self.button.move(70, 110)

        self.button_2 = QPushButton(self)
        self.button_2.setText("Add Data 2")
        self.button_2.adjustSize()
        self.button_2.clicked.connect(self.onButtonClick)
        self.button_2.move(170, 110)

        #create tablewidget that has extended key press event for being able to copy/paste into
        self.table = TableWithCopy(self)
        self.table.setGeometry(10, 150, 350, 380)

        # Font for table.  NOTE: resetting font/font size means you will HAVE to change
        # the rowheight/setrowheight values within the onCountChanged function.  within
        # onCountChanged function, changes row height dynamically based on if there is word wrapping
        # or not, if not, makes the row height smaller to make a more compact table, which is dependent
        # on the font/font size.
        fnt = self.table.font()
        fnt.setPointSize(8)
        self.table.setFont(fnt)

        # set that the row height can be made all the way to 0
        self.table.verticalHeader().setMinimumSectionSize(0)

        # set maximum size the cell can be to prevent really long cells being made, word wrapping
        # dont later with expanded row heights to account for cells with lots of info
        self.table.horizontalHeader().setMaximumSectionSize(100)

        #set headers to be aligned in center of cell
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)

        # alternate row colors for better lookign table
        self.table.setAlternatingRowColors(True)

        # for setting default row height... not really needed, resizeRowsToContents()
        # is called later to resize the heights
        #self.table.verticalHeader().setDefaultSectionSize(17)

        # fonts of the headers and making bold (don't think border-radius in this stylesheet is actually doing anything
        stylesheet = ("QHeaderView::section{font-weight: bold; font-size: 8pt; border-radius:14px;} QTableView::item {padding: 0px;}")

        # set style sheet
        self.table.setStyleSheet(stylesheet)

        # set when cell is clicked it calls function to get value of that cell
        self.table.cellClicked.connect(self.get_item)

        # show window
        self.show()

    def get_item(self, row, column):
        # get cell value
        cell_value = self.table.item(row, column).text()
        print(cell_value)


    def onButtonClick(self):
        # which button clicked for which data to add
        button_clicked = self.sender().text()

        # start Qthread to collect data for table.  on a separate thread so that UI remains responsive
        self.calc = External(button_clicked)
        self.calc.secondcountChanged.connect(self.newCountChanged)
        self.calc.start()

    # take lists passed from Qthread to setup new row(s), column and populate data to the cells
    # the cells/headers and then resize row heights based on content/word wrapping
    def newCountChanged(self, column, header):
        # insert new column to table
        self.table.insertColumn(self.table.columnCount())

        # set header of the column just added, -1 on the columncount() because adding
        # data to columns/rows is done by index location which starts at 0
        self.table.setHorizontalHeaderItem(self.table.columnCount() - 1, QTableWidgetItem(header))

        # if number of rows is larger than current table, add more rows to accomodate
        delta_rows = len(column) - self.table.rowCount()
        while delta_rows > 0:
            self.table.insertRow(self.table.rowCount())
            delta_rows -= 1

        # add data to all the rows from the list.  -1 is on the columncount() because adding
        # data to columns/rows is done by index location which starts at 0
        for i in range(len(column)):
            self.table.setItem(i, self.table.columnCount() - 1, QTableWidgetItem(column[i]))

        # reset row/column heights based on content, needed you can see all the content
        # in a cell if there is word wrapping
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

        # if row height is at the default height of 23, which means there's no wordwrapping going on,
        # then set the row height to 17 to make the table more compact.  NOTE: the default height of 23
        # is DEPENDENT on the FONT SIZE called out at the outset of the UI creation due to using
        # resizeRowsToContents() on the line above.  I don't think there is another way to do this as running
        # resizeRowsToContents() resets row heights based on content, which is needed for word wrapping, but resets
        # row heights to some arbitrary height based on the font for all other rows that don't have word wrapping.

        # you could just get rid of this and just use the resizeRowsToContents() above only, if you don't care
        # about having a more compact table.

        for row in range(self.table.rowCount()):
            if self.table.rowHeight(row) == 23:
                self.table.setRowHeight(row, 17)


class External(QThread):
    # signal for newCountChanged function
    secondcountChanged = pyqtSignal(list, str)

    def __init__(self, button_clicked):
        super().__init__()
        self.button_clicked = button_clicked

    def run(self):

        # collect data for table.  This data set just as an example, can collect data into lists for anything
        column1 = ["aaaaaaaaaaa aaaaaaaa", "b", "c", "d", "a", "b", "c", "d", "a", "b", "c", "end"]
        column2 = ["JOB111", "JOB2", "JOB3333333 333333", "JOB5", "JOB1", "JOB2", "JOB3", "JOB5", "JOB1", "JOB5"]

        column1_header = "header1"
        column2_header = "header2"

        if "1" in self.button_clicked:
            self.secondcountChanged.emit(column1, column1_header)
        if "2" in self.button_clicked:
            self.secondcountChanged.emit(column2, column2_header)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Actions()
    sys.exit(app.exec_())






