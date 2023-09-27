from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import QTableWidget, QLineEdit, QPushButton, QApplication, QMainWindow, QVBoxLayout, QWidget, \
    QTableWidgetItem, QFormLayout, QLabel, QMessageBox, QHeaderView, QAction, QActionGroup, QStatusBar, QMenuBar
from PyQt5.QtCore import Qt, QThread, pyqtSignal


import pandas as pd
import sys
import os
import time

from datetime import date
import win32print
import win32com.client

# can't remember what the pythoncom stuff does, i think it had something to do with when turning it into an exe file
# and for some reason needs to be within the class where win32 commands are done
#import pythoncom
#pythoncom.CoInitialize()

# DIA vendors (possible add this to the UI later at some point with option to modify DIA vendors on the UI)
# for now just making it hard coded
dia_vendors = ["Custom Interface, Inc", "Enviro Systems Incorporated", "Lee Aerospace",
               "McFarlane Aviation", "Precise Flight", "Rapid Manufacturing", "ISCO"]

# get printers for UI
printer_names = ([printer[2] for printer in win32print.EnumPrinters(2)])
default_printer = win32print.GetDefaultPrinter()


class MainWindow(QMainWindow):

    def __init__(self, parent = None):
        super(MainWindow, self).__init__(parent)
        # can't remember what the pythoncom stuff does, i think it had something to do with when turning it into an exe file
        # and for some reason needs to be within the class where win32 commands are done
        import pythoncom
        pythoncom.CoInitialize()

        self.setWindowTitle("Admin QC REC")
        self.setGeometry(200, 200, 1000, 400)

        # add menu with printers for user selection
        self.printer_menu = self.menuBar()
        self.selmenu = self.printer_menu.addMenu("Select Printer")

        # put printer menu items in actiongroup so that only 1 can be checked at a time
        self.printer_action_group = QActionGroup(self)

        # add menu selection for every printer so user can choose
        for i in printer_names:
            menu_item = QAction(i, self)
            menu_item.setCheckable(True)

            # add menu item to menu and action group
            self.selmenu.addAction(menu_item)
            self.printer_action_group.addAction(menu_item)

        self.about_menu = self.printer_menu.addMenu("Info")
        self.about_item = QAction("How It Works")
        self.about_menu.addAction(self.about_item)
        self.about_item.triggered.connect(self.about_info)

        # set action group so that only 1 item can be checked at a time
        self.printer_action_group.setExclusive(True)

        self.form_layout = QFormLayout()
        self.po_supplier_path = QLineEdit()
        self.qc_rec_path = QLineEdit()
        self.form_layout.addRow(QLabel("PO/Supplier File:"), self.po_supplier_path)
        self.form_layout.addRow(QLabel("QC Rec File:"), self.qc_rec_path)

        self.find_data_button = QPushButton("Find Data")
        self.find_data_button.clicked.connect(self.find_data)

        self.warning_label = QLabel("***Note: When getting QC REC file data from Epic Queries, must pick day after for end date!")
        self.warning_label.adjustSize()
        self.update_label = QLabel()

        self.query = QLineEdit()
        self.query.setPlaceholderText("Search...")
        self.query.textChanged.connect(self.search)

        self.table = QTableWidget()
        # make cells stretch to width of UI for better look
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)

        # alternate row colors for better lookign table
        self.table.setAlternatingRowColors(True)

        # set style sheet for selections.  Note:  when overwrriting selection background color
        # text selection color needs to be set or the text color will change automatically
        # when clicking on a cell vs using the search feature
        self.table.setStyleSheet("selection-background-color: cyan; selection-color: black;")

        #fnt = QFont()
        #fnt.setBold(True)

        self.column_headers = ["Part No.", "PO No.", "Lot/Serial No.", "Rec Date", "Supplier",
                               "FAI Needed", "PO Button", "FAI Button", "Open Parts Folder"]
        self.table.setColumnCount(len(self.column_headers))
        self.table.setHorizontalHeaderLabels(self.column_headers)
        #self.table.setFont(fnt)

        layout = QVBoxLayout()

        layout.addLayout(self.form_layout)
        layout.addWidget(self.warning_label, alignment=Qt.AlignHCenter)
        layout.addWidget(self.find_data_button, alignment=Qt.AlignHCenter, stretch=False)
        layout.addWidget(self.update_label, alignment=Qt.AlignHCenter)
        layout.addWidget(self.query)
        layout.addWidget(self.table)

        self.table.resizeRowsToContents()
        #self.table.resizeColumnsToContents()

        # use status bar for program updates instead of self.update_label
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)

        w = QWidget()
        w.setLayout(layout)
        self.setCentralWidget(w)

    def about_info(self):
        msg = QMessageBox()
        msg.setText("<u>How This Works:</u>"
                    "<p>    - If part and/or supplier not on PO/Supplier list - Returns need FAI</p>"
                    "<p>    - If on there and PO/REC match, in theory this should mean it needs an FAI</p>"
                    "<p>    - If different PO# and/or REC date, in theory, should not need an FAI</p>"
                    "<p>    - If NCR checkmarked or is a PK99, this may or may not need an FAI, asks you to check manually</p>"
                    "<p>    - If more than 2 years since last inspected one, needs FAI (unless PK99, then asks you to check manually</p>")
        msg.setWindowTitle("Info")
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()

    def search(self, s):
        if s == "":
            self.table.clearSelection()
        else:
            items = self.table.findItems(s, Qt.MatchContains)
            if items:  # we have found something
                item = items[0]  # take the first
                self.table.setCurrentItem(item)

                # set items in the first 5 columns to be selected when doing a search
                # don't want 6th column as i don't want to overrite the background color
                for i in range(5):
                    self.table.item(item.row(), i).setSelected(True)

    def this_button(self):

        # po scans folder
        purchasing_scans_path = "\\\\NAS3\\Canon Scanner\\Purchasing Scans\\"

        # get which button was clicked
        sender = self.sender()

        # no texactly sure how this works, but is needed to get the object name of the pushbutton
        push_button = self.findChild(QPushButton, sender.objectName())

        # get row data based on which button was clicked
        row_data = []

        for col in range(self.table.columnCount()):
            item = self.table.item(int(push_button.objectName()), col)
            # gets only text from the rows
            if item != None:
                row_data.append(item.text())

        # if OPen PO button clicked, find files with partnumber in it and find matching receive dates and open them
        if "Open" in sender.text():
            files_paths = []

            # append all files that have PO and same date
            for i in self.po_folder_files:
                # if po number in readfolder
                if row_data[1] in i:
                    # get full path of file
                    po_path = purchasing_scans_path + i

                    # get modified time of file
                    ti_m = os.path.getmtime(po_path)
                    m_ti = time.ctime(ti_m)

                    # this to be able to convert
                    t_obj = time.strptime(m_ti)

                    # transform time object to time stam
                    T_stamp = time.strftime("%Y-%m-%d", t_obj)

                    # if file modified time stamp is same as day im looking for append file path for opening
                    if row_data[3] in T_stamp:
                        files_paths.append(po_path)

            # check if any file found, if not display error message, if found, open all of them
            if len(files_paths) == 0:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Couldn't find PO PDF that matches receive date")
                msg.setWindowTitle("No File")
                msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
                retval = msg.exec_()

            else:
                for i in files_paths:
                    os.startfile(i)

        # find excel path to fai and then print it
        if "Print" in sender.text():
            partnumber = row_data[0]

            fai_excel, faipartpath = self.find_fai_path(partnumber)

            if os.path.exists(fai_excel):
                self.print_excel(fai_excel)
            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Couldn't find FAI.  Could be not made or a PROTO part!")
                msg.setWindowTitle("No File")
                msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
                retval = msg.exec_()

        # open windows explorer to parts folder
        if "Folder" in sender.text():
            partnumber = row_data[0]
            fai_excel, faipartpath = self.find_fai_path(partnumber)

            if os.path.exists(faipartpath):
                os.startfile(faipartpath)
            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Couldn't Parts folder.  Could be not made or a PROTO part!")
                msg.setWindowTitle("No Folder")
                msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
                retval = msg.exec_()


    def find_fai_path(self, part_number):
        # give initial value to returned value
        fai_excel = "No"

        partnumber = part_number
        part = partnumber[0:10]
        chapterfai = partnumber[2:6]
        upperchapter = partnumber[2:4]

        if chapterfai == "2710":
            chapterfai = "2710 Aileron"
        if chapterfai == "2720":
            chapterfai = "2720 rudder control sys"
        if chapterfai == "2730":
            chapterfai = "2730 elevator control sys"
        if chapterfai == "2750":
            chapterfai = "2750 flaps"
        if chapterfai == "2820":
            chapterfai = "2820 wing"
        if chapterfai == "3200":
            chapterfai = "3200 landing gear"
        if chapterfai == "3210":
            chapterfai = "3210 landing gear"
        if chapterfai == "3220":
            chapterfai = "3220 landing gear"
        if chapterfai == "3230":
            chapterfai = "3230 landing gear"
        if chapterfai == "3240":
            chapterfai = "3240 wheels, brakes"
        if chapterfai == "3250":
            chapterfai = "3250 steering damper"
        if chapterfai == "5210":
            chapterfai = "5210 doors"
        if chapterfai == "5310":
            chapterfai = "5310 assy fuselage bond"
        if chapterfai == "5320":
            chapterfai = "5320 floors"
        if chapterfai == "5330":
            chapterfai = "5330 floors"
        if chapterfai == "5510":
            chapterfai = "5510 horizontal"
        if chapterfai == "5520":
            chapterfai = "5520 elevator"
        if chapterfai == "5540":
            chapterfai = "5540 rudder"
        if chapterfai == "5700":
            chapterfai = "5700 wing bonded assy"
        if chapterfai == "5710":
            chapterfai = "5710 wing"
        if chapterfai == "5740":
            chapterfai = "5740 wing"
        if chapterfai == "5750":
            chapterfai = "5750 wing"
        if chapterfai == "7120":
            chapterfai = "7120 powerplant"

        folderfai = "O:\\Quality\\Quality Documents\\Inspection Reports\\First Article Inspection Reports\\"

        faipartpath = '\\'.join([folderfai, upperchapter, chapterfai, part])

        if upperchapter == "00":
            faipartpath = '\\'.join([folderfai, upperchapter, part])

        # see if path exists, if so iterate through and find the excel
        if os.path.exists(faipartpath):
            readfolder = os.listdir(faipartpath)

            for i in readfolder:
                excel_filename = part_number + ".xlsx"
                if excel_filename in i:
                    fai_excel = faipartpath + "\\" + excel_filename

        return fai_excel, faipartpath


    def print_excel(self, excel_path):
        # initialize printer value
        printer = "None"


        # get which printer is checkmarked, if none send message printer not chosen
        for i in self.printer_action_group.actions():
            if i.isChecked():
                printer = i.text()

        # print excel FAI, set all sheets to fit to page, reset default printer back to original
        if "None" not in printer:
            win32print.SetDefaultPrinter(printer)

            # use Excel API to set all pages to fit to page before printing
            o = win32com.client.Dispatch('Excel.Application')
            o.visible = False
            wb = o.Workbooks.Open(excel_path)

            for sh in wb.Sheets:
                sh.PageSetup.Zoom = False
                sh.PageSetup.FitToPagesTall = 1
                sh.PageSetup.FitToPagesWide = 1

            wb.PrintOut()
            wb.Close()

            win32print.SetDefaultPrinter(default_printer)
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Pick a Printer first!")
            msg.setWindowTitle("Select Printer")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            retval = msg.exec_()

    def find_data(self):
        self.statusbar.showMessage("Reading PO Scans......")

        self.start_data = Data_retrieval(self.po_supplier_path.text(), self.qc_rec_path.text())
        self.start_data.update_tableChanged.connect(self.on_updatetable)
        self.start_data.update_labelChanged.connect(self.on_updateChanged)
        self.start_data.po_scans_folder.connect(self.po_folder_data)
        self.start_data.start()

    # for files in po folder (this is so "open po" button does scan dir everytime)
    def po_folder_data(self, po_folder):
        self.po_folder_files = po_folder

    def on_updateChanged(self, label):

        # not using self update-label anymore, using self.statusbar for progrma updates
        #self.update_label.setText(label)
        #self.update_label.adjustSize()

        self.statusbar.showMessage(label)

    def on_updatetable(self, table_list):
        # clear table to start if find button hit again to re-populate table
        self.table.clear()

        # set size of column and headers
        self.table.setColumnCount(len(self.column_headers))
        self.table.setRowCount(len(table_list))
        self.table.setHorizontalHeaderLabels(self.column_headers)

        # populate table
        for c in range(0, len(self.column_headers)):
            for r in range(0, len(table_list)):
                if c < 6:
                    s = table_list[r][c]
                    i = QTableWidgetItem(s)

                    # change background color of column 5
                    if c == 5:
                        if "No" in s:
                            i.setBackground(QColor('green'))
                        if "Print" in s:
                            i.setBackground(QColor('yellow'))
                        if "NCR" in s:
                            i.setBackground(QColor('orange'))
                        if "Possible" in s:
                            i.setBackground(QColor("orange"))

                    # use below line if you want editable cells
                    # i.setFlags(i.flags() ^ Qt.ItemIsEditable)
                    self.table.setItem(r, c, i)
                if c == 6:
                    po_btn = QPushButton("Open PO")
                    po_btn.setObjectName(str(r))
                    po_btn.clicked.connect(self.this_button)
                    self.table.setCellWidget(r, c, po_btn)
                if c == 7:
                    fai_btn = QPushButton("Print FAI")
                    fai_btn.setObjectName(str(r))
                    fai_btn.clicked.connect(self.this_button)
                    self.table.setCellWidget(r, c, fai_btn)
                if c == 8:
                    folder_btn = QPushButton("Parts Folder")
                    folder_btn.setObjectName(str(r))
                    folder_btn.clicked.connect(self.this_button)
                    self.table.setCellWidget(r, c, folder_btn)

        # set update text
        self.statusbar.showMessage("Done.")

        # resize contents for better fit
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()


class Data_retrieval(QThread):
    update_labelChanged = pyqtSignal(str)
    update_tableChanged = pyqtSignal(list)
    po_scans_folder = pyqtSignal(list)

    def __init__(self, po_path, qc_path):
        super(Data_retrieval, self).__init__()

        # remove any quotes from string path
        removed_quotes_po = po_path.replace('"', "")
        removed_quotes_qc = qc_path.replace('"', "")

        self.po_path = removed_quotes_po
        self.qc_path = removed_quotes_qc

    # iterate through QC rec excel file
    def run(self):

        # po scans folder
        purchasing_scans_path = "\\\\NAS3\\Canon Scanner\\Purchasing Scans\\"

        # get files from dir, doing scanning of po folder in the thread and passing to mainwindow
        # otherwise would have to scan this dir everytime "open po" which takes 6 seconds everytime
        po_readfolder = os.listdir(purchasing_scans_path)

        table_list = []
        if os.path.exists(self.po_path) and os.path.exists(self.qc_path):

            self.update_labelChanged.emit("Reading Data......")

            po_df = pd.read_excel(self.po_path)
            qc_df = pd.read_excel(self.qc_path)

            # iterate through QC dataframe
            for index in range(len(qc_df)):

                # filter for AKs/PKs/HKs only
                if "AK" in str(qc_df.iloc[index, 0]) or "PK" in str(qc_df.iloc[index, 0]) or "HK" in str(qc_df.iloc[index, 0]):

                    part_number = str(qc_df.iloc[index, 0][0:12])

                    # get rec date and split just to get date and remove time
                    raw_rec_date = str(qc_df.iloc[index, 1])
                    rec_date = raw_rec_date.split()[0]

                    # get lot number from whichever column
                    if "nan" in str(qc_df.iloc[index, 7]):
                        lot_number = str(qc_df.iloc[index, 6])
                    else:
                        lot_number = str(qc_df.iloc[index, 7])

                    sup_name = str(qc_df.iloc[index, 9])
                    po_number = str(qc_df.iloc[index, 4].split()[-1])

                    # activate function to reference po excel for check if FAI is needed or not
                    fai_needed = self.check_fai(po_df, part_number, sup_name, po_number, rec_date)

                    temp_list = [part_number, po_number, lot_number, rec_date, sup_name, fai_needed]
                    table_list.append(temp_list)

        # sort table list before emitting to UI
        table_list.sort(key=lambda x: x[0])

        # emit to UI
        self.update_tableChanged.emit(table_list)
        self.po_scans_folder.emit(po_readfolder)


    # function to determine whether FAI is needed and whether it's a DIA vendor
    def check_fai(self, po_df, part_number, sup_name, po_number, rec_date):
        # initialize return value
        fai_needed = ""

        # determine if part number is even on list first off, if so, get all index rows for number of times part appears on list
        part_appearances = []
        for index in range(len(po_df)):
            # if part number and supplier name match add to index for iteration further down
            if part_number.upper() in str(po_df.iloc[index, 0]).upper() and \
                    sup_name.upper() in str(po_df.iloc[index, 3]).upper():
                part_appearances.append(index)

        # if not found in list PRINT FAI
        if len(part_appearances) == 0:
            # convert all to upper comparison and compare
            if sup_name.upper() in [w.upper() for w in dia_vendors]:
                fai_needed = "Print Vendor FAI"
            else:
                fai_needed = "Print FAI"

        # for how many times part appears on list get total vendors/pos for comparison later
        else:
            if len(part_appearances) > 0:
                #print(part_number, po_number, rec_date)
                total_vendors = []
                total_pos = []
                date_last_rec = []
                fai_completed = []
                current_rec_date = []
                ncr = []
                for i in part_appearances:
                    total_vendors.append(str(po_df.iloc[i, 3]))
                    total_pos.append(str(po_df.iloc[i, 9]))

                    # get date, split out to remove time
                    raw_date_last_rec = str(po_df.iloc[i, 10])
                    date_last_rec.append(raw_date_last_rec.split()[0])

                    fai_completed.append(str(po_df.iloc[i, 6]))

                    # get date, split out to remove time
                    raw_current_rec_date = str(po_df.iloc[i, 4])
                    current_rec_date.append(raw_current_rec_date.split()[0])

                    # pass if ncr checkmarked
                    ncr.append(str(po_df.iloc[i, 5]))

                # if vendor received not in vendor list, print fai or vendor fai
                if sup_name.upper() not in [w.upper() for w in total_vendors]:
                    if sup_name.upper() in [k.upper() for k in dia_vendors]:
                        fai_needed = "Print Vendor FAI"
                    else:
                        fai_needed = "Print FAI"

                 # if supplier received IS in vendor list
                else:
                   # if true in ncr column, it might need an FAI
                    if "True" in ncr:
                        fai_needed = "Possible NCR FAI? Check Manually, ref. QPM-008"
                    else:
                        # if fai_completed marks as true fai, not needed, this is to weed out some parts
                        # that we receive on the same po over and over again
                        if "True" in fai_completed:
                            fai_needed = "No"

                        # if fai completed not marked
                        else:
                            # if any values in total_pos not in value, then fai not needed
                            # this is an error check if part/rev appears more than once on list
                            # but with different POs, due to things like color codes , which we don't need an FAI for
                            if any(po_number not in x for x in total_pos):
                                fai_needed = "No"
                            # if supplier received in vendor list and PO matches 1st one received,  need vendor FAI or FAI
                            else:
                                # if there's a date that doens't match our rec date, fai not needed
                                # this is to rule out HK's or AK's received on different date later but with same PO#
                                if any(rec_date not in x for x in current_rec_date):
                                    fai_needed = "No"
                                else:
                                    if sup_name.upper() in [k.upper() for k in dia_vendors]:
                                        fai_needed = "Print Vendor FAI"
                                    else:
                                        fai_needed = "Print FAI"

                    # nat/nan error check to make sure theres an actual date value returned from cell
                    if "NaT" not in date_last_rec and "nan" not in date_last_rec and len(date_last_rec) != 0:
                        # loops thru all rec dates to make sure at least one rec less than 730 days ago
                        time_differences = []

                        for i in date_last_rec:

                        # check for how many days elapse, if elapsed days > 730  days (2 years) , needs FAI, this overrides
                        # everything else
                            curr_time = time.strftime("%Y-%m-%d", time.localtime())

                            t_obj = time.strptime(i, "%Y-%m-%d")
                            t_obj1 = time.strptime(curr_time, "%Y-%m-%d")

                            time_obj_old = time.mktime(t_obj)
                            time_obj_new = time.mktime(t_obj1)

                            delta = time_obj_new - time_obj_old
                            time_differences.append(int(delta / 86400))

                        # if all last inspections are greater than 2 years
                        if all(y > 730 for y in time_differences):
                            # add check if PK99 greater than 2 years, special case
                            # where the inspection may have been done on job instead
                            if "PK99" in part_number:
                                fai_needed = "FAI Possible, Manually Check this, last inspect maybe on JOB instead of PO"
                            else:
                                if sup_name.upper() in [k.upper() for k in dia_vendors]:
                                    fai_needed = "Print Vendor FAI"
                                else:
                                    fai_needed = "Print FAI"
                            #print(part_number)
                            #print("RECEIVE DATE " + ' '.join(current_rec_date))
                            #print("DATE LAST REC " + ' '.join(date_last_rec))

                # if RMA in po_number, this should auto be assumed as not needing an FAI,
                # as RMA means we had already received it once, and we should hae done an FAI
                # at that point in time, if required
                if "RMA" in po_number:
                    fai_needed = "No"

        return fai_needed

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
