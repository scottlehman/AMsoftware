import sys
import os
import re
import gzip
import webbrowser
from tinydb import TinyDB, Query, where
from PyQt5.QtWidgets import QMainWindow, QComboBox, QLabel, QPushButton, QWidget, QApplication, QLineEdit, QMessageBox, QPlainTextEdit, QTableWidget, QListWidget, QTableWidgetItem
servicedesk = TinyDB(R'storage\JSON\servicedesk.json')
lawson = TinyDB(R'storage\JSON\lawson.json')

class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Archived Service Desk & Lawson Orders')
        self.setGeometry(550, 150, 1030, 850)
        self.combo_lbl = QLabel(self)
        self.combo_lbl.setText('ServiceDesk:')
        self.combo_lbl.move(10, 10)
        self.combo = QComboBox(self)
        self.combo.addItem('CO Number')
        self.combo.addItem('Affected User')
        self.combo.addItem('Requestor')
        self.combo.addItem('Summary')
        self.combo.addItem('Description')
        self.combo.move(10, 35)
        self.search_lbl = QLabel(self)
        self.search_lbl.setText('Keywords:')
        self.search_lbl.move(120, 10)
        self.search = QLineEdit(self)
        self.search.move(120, 35)   
        self.btn = QPushButton('Search', self)
        self.btn.resize(80, 30)
        self.btn.move(224, 35)
        self.btn.clicked.connect(self.search_servicedesk)
        self.po_btn = QPushButton('Search Lawson POs', self)
        self.po_btn.resize(150, 30)
        self.po_btn.move(305, 35)
        self.po_btn.clicked.connect(self.search_podb)
        self.attach_btn = QPushButton('Attachments Folder', self)
        self.attach_btn.resize(120, 30)
        self.attach_btn.move(561, 390)
        self.attach_btn.clicked.connect(self.open_repo)
        self.co_numb = QLabel(self)
        self.co_numb.setText('CO Number:')
        self.co_numb.move(10, 60)
        self.co_numb_value = QPlainTextEdit(self)
        self.co_numb_value.setPlaceholderText('')
        self.co_numb_value.setReadOnly(True)
        self.co_numb_value.move(10, 85)
        self.affected_user = QLabel(self)
        self.affected_user.setText('Affected User:')
        self.affected_user.move(122, 60)
        self.affected_user_value = QPlainTextEdit(self)
        self.affected_user_value.resize(180, 30)
        self.affected_user_value.setPlaceholderText('')
        self.affected_user_value.setReadOnly(True)
        self.affected_user_value.move(122, 85)
        self.requestor_lbl = QLabel(self)
        self.requestor_lbl.setText('Requestor:')
        self.requestor_lbl.move(315, 60)
        self.requestor_value = QPlainTextEdit(self)
        self.requestor_value.resize(180, 30)
        self.requestor_value.setPlaceholderText('')
        self.requestor_value.setReadOnly(True)
        self.requestor_value.move(315, 85)
        self.summary = QLabel(self)
        self.summary.setText('Summary:')
        self.summary.move(360, 115)
        self.summary_value = QPlainTextEdit(self)
        self.summary_value.setPlaceholderText('')
        self.summary_value.resize(320, 80)
        self.summary_value.setReadOnly(True)
        self.summary_value.move(360, 140)
        self.description = QLabel(self)
        self.description.setText('Description:')
        self.description.move(10, 115)
        self.description_value = QPlainTextEdit(self)
        self.description_value.resize(316, 305)
        self.description_value.setPlaceholderText('')
        self.description_value.setReadOnly(True)
        self.description_value.move(10, 140)
        self.attmnts_list_lbl = QLabel(self)
        self.attmnts_list_lbl.setText('Attachments:')
        self.attmnts_list_lbl.move(360, 215)
        self.attmnts_list = QListWidget(self)
        self.attmnts_list.resize(320, 140)
        self.attmnts_list.move(360, 240)
        self.attmnts_list.stackUnder(self.attach_btn)
        self.attmnts_list.itemClicked.connect(self.unzip)
        self.actlog_lbl = QLabel(self)
        self.actlog_lbl.setText('Activity Log:')
        self.actlog_lbl.move(10, 445)
        self.actlog = QTableWidget(self)
        self.actlog.move(10, 475)
        self.actlog.resize(675, 360)
        self.actlog.verticalHeader().sectionDoubleClicked.connect(self.display_row)
        self.comment_lbl = QLabel(self)
        self.comment_lbl.setText('Activity Log Details:')
        self.comment_lbl.move(700, 445)
        self.comments = QPlainTextEdit(self)
        self.comments.setReadOnly(True)
        self.comments.move(700, 475)
        self.comments.resize(320, 360)
        self.lawson_combo_lbl = QLabel(self)
        self.lawson_combo_lbl.setText('Lawson:')
        self.lawson_combo_lbl.move(700, 10)
        self.lawson_combo = QComboBox(self)
        self.lawson_combo.addItem('PO Number')
        self.lawson_combo.addItem('CO Number')
        self.lawson_combo.addItem('Vendor')
        self.lawson_combo.addItem('Item')
        self.lawson_combo.addItem('Description')
        self.lawson_combo.addItem('Cost Center')
        self.lawson_combo.addItem('BRN Number')
        self.lawson_combo.move(700, 35)
        self.lawson_search_lbl = QLabel(self)
        self.lawson_search_lbl.setText('Keywords:')
        self.lawson_search_lbl.move(810, 10)
        self.lawson_search = QLineEdit(self)
        self.lawson_search.move(810, 35)
        self.lawson_btn = QPushButton('Search', self)
        self.lawson_btn.resize(80, 30)
        self.lawson_btn.move(914, 35)
        self.lawson_btn.clicked.connect(self.search_lawson)
        self.lawson_list_lbl = QLabel(self)
        self.lawson_list_lbl.setText('Purchase Order:')
        self.lawson_list_lbl.move(700, 115)
        self.lawson_list = QPlainTextEdit(self)
        self.lawson_list.setReadOnly(True)
        self.lawson_list.resize(320, 305)
        self.lawson_list.move(700, 140)
        self.show()

    def search_servicedesk(self):
        global q
        if self.search.text() == '':
            self.error_message()
            return
        elif self.combo.currentText() == 'CO Number':
            q = servicedesk.search(where('co').matches(F'{self.search.text()}.*'))
        elif self.combo.currentText() == 'Affected User':
            q = servicedesk.search(where('endUser').search(F'{self.search.text()}+', flags=re.IGNORECASE))
        elif self.combo.currentText() == 'Requestor':
            q = servicedesk.search(where('req').search(F'{self.search.text()}+', flags=re.IGNORECASE))
        elif self.combo.currentText() == 'Summary':
            q = servicedesk.search(Query().summ.search(F'{self.search.text()}+', flags=re.IGNORECASE))
        elif self.combo.currentText() == 'Description':
            q = servicedesk.search(Query().desc.search(F'{self.search.text()}+', flags=re.IGNORECASE))
        if len(q) == 0:
            self.error_message()
            return
        elif len(q) == 1: 
            self.populate_main(q[0])
        elif len(q) > 1:
            self.populate_new_list_window(q)

    def error_message(self):
        msg = QMessageBox()
        msg.setWindowTitle('Missing fields required.')
        msg.setIcon(QMessageBox.Information)
        msg.setText('No results or Missing required input fields.')
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def populate_main(self, c):
        self.co_numb_value.setPlainText('')
        self.affected_user_value.setPlainText('')
        self.description_value.setPlainText('')
        self.summary_value.setPlainText('')
        self.requestor_value.setPlainText('')
        self.actlog.setRowCount(0)
        self.comments.setPlainText('')
        self.lawson_list.setPlainText('')
        self.attmnts_list.clear()
        self.co_numb_value.setPlainText(c['co'])
        self.affected_user_value.setPlainText(c['endUser'])
        self.requestor_value.setPlainText(c['req'])
        self.summary_value.setPlainText(c['summ'])
        self.description_value.setPlainText(c['desc'])
        self.actlog.setColumnCount(5)
        self.actlog.setHorizontalHeaderLabels(['Type', 'Date', 'Analyst', 'Description', 'Comments/Details'])
        for i in c['actLog']:
            rowPosition = self.actlog.rowCount()
            self.actlog.insertRow(rowPosition)
            self.actlog.setItem(rowPosition, 0, QTableWidgetItem(i['type']))
            self.actlog.resizeColumnToContents(0) 
            self.actlog.setItem(rowPosition, 1, QTableWidgetItem(i['createdOn']))
            self.actlog.resizeColumnToContents(1) 
            self.actlog.setItem(rowPosition, 2, QTableWidgetItem(i['createdBy']))
            self.actlog.resizeColumnToContents(2) 
            self.actlog.setItem(rowPosition, 3, QTableWidgetItem(i['descrip'])) 
            self.actlog.resizeColumnToContents(3)
            self.actlog.setItem(rowPosition, 4, QTableWidgetItem(i['details'])) 
            self.actlog.resizeColumnToContents(4)       
        for attmnt in c['attmnts']:
            self.attmnts_list.addItem('{}'.format(attmnt['ogName']))
                     
    def search_podb(self):
        global l
        l = lawson.search(where('CO_Number') == self.co_numb_value.toPlainText())
        if self.co_numb_value.toPlainText() == '' or len(l) == 0:
            self.error_message()
            return
        elif len(l) == 1:
            self.one_lawson_result(l)
        elif len(l) > 1:
            self.multi_lawson_result(l)
        
    def populate_new_list_window(self, r):
        self.sw = new_list_window()
        self.sw.verticalHeader().sectionDoubleClicked.connect(self.get_info)
        for i in r:
            rowPosition = self.sw.rowCount()
            self.sw.insertRow(rowPosition)
            self.sw.setItem(rowPosition, 0, QTableWidgetItem(i['actLog'][0]['createdOn']))
            self.sw.resizeColumnToContents(0)
            self.sw.setItem(rowPosition, 1, QTableWidgetItem(i['co']))
            self.sw.resizeColumnToContents(1)
            self.sw.setItem(rowPosition, 2, QTableWidgetItem(i['endUser']))
            self.sw.resizeColumnToContents(2)
            self.sw.setItem(rowPosition, 3, QTableWidgetItem(i['summ']))
            self.sw.resizeColumnToContents(3)
            self.sw.setItem(rowPosition, 4, QTableWidgetItem(i['desc']))
            self.sw.resizeColumnToContents(4)
        self.sw.show()

    def get_info(self, row):
        for x in q:
            if self.sw.item(row, 1).text() == x['co']:
                self.populate_main(x)

    def display_row(self, t):
        self.comments.setPlainText('')
        self.comments.appendPlainText(self.actlog.item(t, 0).text())
        self.comments.appendPlainText(self.actlog.item(t, 1).text())
        self.comments.appendPlainText(self.actlog.item(t, 2).text())
        self.comments.appendPlainText(self.actlog.item(t, 3).text())
        self.comments.appendPlainText(self.actlog.item(t, 4).text())

    def unzip(self, i):
        if i.text().startswith('https'):
            webbrowser.open_new_tab(F'{i.text()}')
        for co in q:
            if self.co_numb_value.toPlainText() == co['co']:
                for attch in co['attmnts']:
                    if i.text() == attch['ogName']:
                        source = R'storage\GZ\{}'.format(attch['zipName'])
                        dest = R'Attachments\{}'.format(attch['ogName'])
                try:
                    with gzip.open(source, 'rb') as s_file, open(dest, 'wb') as d_file:
                        while True:
                            block = s_file.read(65536)
                            if not block:
                                break
                            else:
                                d_file.write(block)
                        d_file.write(block)
                    os.startfile(dest)
                except (PermissionError, FileNotFoundError) as err:
                    print(err)
                    pass

    def open_repo(self):
        os.startfile(R'Attachments')

    def search_lawson(self):
        if self.lawson_search.text() == '':
            self.error_message()
            return
        elif self.lawson_combo.currentText() == 'PO Number':
            p = lawson.search(where('PO_NUMBER').matches(F'{self.lawson_search.text()}.*')) 
        elif self.lawson_combo.currentText() == 'CO Number':
            p = lawson.search(where('CO_Number').matches(F'{self.lawson_search.text()}.*'))
        elif self.lawson_combo.currentText() == 'Vendor':
            p = lawson.search(Query().VENDOR_NAME.search(F'{self.lawson_search.text()}+', flags=re.IGNORECASE))
        elif self.lawson_combo.currentText() == 'Item':
            p = lawson.search(Query().ITEM.search(F'{self.lawson_search.text()}+', flags=re.IGNORECASE))
        elif self.lawson_combo.currentText() == 'Description':
            p = lawson.search(Query().DESCRIPTION.search(F'{self.lawson_search.text()}+', flags=re.IGNORECASE))
        elif self.lawson_combo.currentText() == 'Cost Center':
            p = lawson.search(where('COST_CENTER') == F'{self.lawson_search.text()}')
        elif self.lawson_combo.currentText() == 'BRN Number':
            p = lawson.search(where('BRN_NUMBER') == F'{self.lawson_search.text()}')
        if len(p) == 0:
            self.error_message()
            return
        elif len(p) == 1:
            self.one_lawson_result(p)
        elif len(p) > 1:
            self.multi_lawson_result(p)

    def one_lawson_result(self, l):
        self.lawson_list.setPlainText('')
        for i in l:
            self.lawson_list.appendPlainText('PO: {}'.format(i['PO_NUMBER']))
            self.lawson_list.appendPlainText('Received: {}'.format(i['RECEIVEDQTY']))
            self.lawson_list.appendPlainText('Cost: {}'.format(i['COST']))
            self.lawson_list.appendPlainText('Date: {}'.format(i['PO_DATE']))
            self.lawson_list.appendPlainText('Cancelled: {}'.format(i['CANCELLED']))
            self.lawson_list.appendPlainText('Requestor: {}'.format(i['REQUESTER']))
            self.lawson_list.appendPlainText('REQ Number: {}'.format(i['REQ_NUMBER']))
            self.lawson_list.appendPlainText('Buyer: {}'.format(i['BUYER']))
            self.lawson_list.appendPlainText('Closed: {}'.format(i['CLOSED']))
            self.lawson_list.appendPlainText('Location: {}'.format(i['LOCATION']))
            self.lawson_list.appendPlainText('Vendor: {}'.format(i['VENDOR_NAME']))
            self.lawson_list.appendPlainText('CO: {}'.format(i['CO_Number']))
            self.lawson_list.appendPlainText('Description: {}'.format(i['DESCRIPTION']))
            self.lawson_list.appendPlainText('Unit: {}'.format(i['UNIT']))
            self.lawson_list.appendPlainText('Item: {}'.format(i['ITEM']))
            self.lawson_list.appendPlainText('Line: {}'.format(i['LINE']))
            self.lawson_list.appendPlainText('Quantity: {}'.format(i['QUANTITY']))
            self.lawson_list.appendPlainText('Account Code: {}'.format(i['ACCOUNT_CODE']))
            self.lawson_list.appendPlainText('Category: {}'.format(i['CATEGORY']))
            self.lawson_list.appendPlainText('Cost Center: {}'.format(i['COST_CENTER']))
            self.lawson_list.appendPlainText('BRN Number: {}'.format(i['BRN_NUMBER']))
            self.lawson_list.appendPlainText('Contract: {}'.format(i['CONTRACT']))
        
    def lawsonPopMain(self, row):
        s = servicedesk.search(where('co').matches(F'{self.lw.item(row, 11).text()}.*'))
        if len(s) == 0:
            self.error_message()
            return
        elif len(s) == 1:
            self.populate_main(s[0])

    def multi_lawson_result(self, l):
        self.lw = LawsonWindow()
        self.lw.verticalHeader().sectionDoubleClicked.connect(self.lawsonPopMain)
        for i in l:
            rowPosition = self.lw.rowCount()
            self.lw.insertRow(rowPosition)
            self.lw.setItem(rowPosition, 0, QTableWidgetItem(i['PO_NUMBER']))
            self.lw.resizeColumnToContents(0)
            self.lw.setItem(rowPosition, 1, QTableWidgetItem(i['RECEIVEDQTY']))
            self.lw.resizeColumnToContents(1)
            self.lw.setItem(rowPosition, 2, QTableWidgetItem(i['COST']))
            self.lw.resizeColumnToContents(2)
            self.lw.setItem(rowPosition, 3, QTableWidgetItem(i['PO_DATE']))
            self.lw.resizeColumnToContents(3)
            self.lw.setItem(rowPosition, 4, QTableWidgetItem(i['CANCELLED']))
            self.lw.resizeColumnToContents(4)
            self.lw.setItem(rowPosition, 5, QTableWidgetItem(i['REQUESTER']))
            self.lw.resizeColumnToContents(5)
            self.lw.setItem(rowPosition, 6, QTableWidgetItem(i['REQ_NUMBER']))
            self.lw.resizeColumnToContents(6)
            self.lw.setItem(rowPosition, 7, QTableWidgetItem(i['BUYER']))
            self.lw.resizeColumnToContents(7)
            self.lw.setItem(rowPosition, 8, QTableWidgetItem(i['CLOSED']))
            self.lw.resizeColumnToContents(8)
            self.lw.setItem(rowPosition, 9, QTableWidgetItem(i['LOCATION']))
            self.lw.resizeColumnToContents(9)
            self.lw.setItem(rowPosition, 10, QTableWidgetItem(i['VENDOR_NAME']))
            self.lw.resizeColumnToContents(10)
            self.lw.setItem(rowPosition, 11, QTableWidgetItem(i['CO_Number']))
            self.lw.resizeColumnToContents(11)
            self.lw.setItem(rowPosition, 12, QTableWidgetItem(i['DESCRIPTION']))
            self.lw.resizeColumnToContents(12)
            self.lw.setItem(rowPosition, 13, QTableWidgetItem(i['UNIT']))
            self.lw.resizeColumnToContents(13)
            self.lw.setItem(rowPosition, 14, QTableWidgetItem(i['ITEM']))
            self.lw.resizeColumnToContents(14)
            self.lw.setItem(rowPosition, 15, QTableWidgetItem(i['LINE']))
            self.lw.resizeColumnToContents(15)
            self.lw.setItem(rowPosition, 16, QTableWidgetItem(i['QUANTITY']))
            self.lw.resizeColumnToContents(16)
            self.lw.setItem(rowPosition, 17, QTableWidgetItem(i['ACCOUNT_CODE']))
            self.lw.resizeColumnToContents(17)
            self.lw.setItem(rowPosition, 18, QTableWidgetItem(i['CATEGORY']))
            self.lw.resizeColumnToContents(18)
            self.lw.setItem(rowPosition, 19, QTableWidgetItem(i['COST_CENTER']))
            self.lw.resizeColumnToContents(19)
            self.lw.setItem(rowPosition, 20, QTableWidgetItem(i['BRN_NUMBER']))
            self.lw.resizeColumnToContents(20)
            self.lw.setItem(rowPosition, 21, QTableWidgetItem(i['CONTRACT']))
            self.lw.resizeColumnToContents(21)
        self.lw.show()

class LawsonWindow(QTableWidget):
    def __init__(self):
        super(LawsonWindow, self).__init__()
        self.setUI()
    
    def setUI(self):
        self.setWindowTitle('Lawson Purchase Orders')
        self.setGeometry(650, 250, 800, 500)
        self.setColumnCount(22)
        self.setHorizontalHeaderLabels(['PO Number', 'Received Qty', 'Cost', 'PO Date', 'Cancelled', 'Requestor', 'Request Number', 'Buyer', 'Closed', 'Location', 'Vendor', 'CO Number', 'Description', 'Unit', 'Item', 'Line', 'Quantity', 'Account Code', 'Category', 'Cost Center', 'BRN Number', 'Contract'])

class new_list_window(QTableWidget):

    def __init__(self):
        super(new_list_window, self).__init__()
        self.setUI()

    def setUI(self):
        self.setWindowTitle('Asset Management TinyDB: Search Results')
        self.setGeometry(650, 250, 800, 500)
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels(['Open Date', 'CO Number', 'Affected End User', 'Summary', 'Description'])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())
