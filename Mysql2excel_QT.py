#coding=utf-8
from PyQt5.QtWidgets import *
import sys
import pymysql
import pandas as pd

class LoginDlg(QDialog):
    def __init__(self, parent=None):
        super(LoginDlg, self).__init__(parent)
        ip = QLabel("Mysql IP：")
        usr = QLabel("账号：")
        pwd = QLabel("密码：")
        port = QLabel("端口：")
        database = QLabel("库名：")
        excel = QLabel("目标CSV：")
        targetexcel = QLabel("导出EXCEL：")
        sql = QLabel("SQL：\r\n(请用\r\n####\r\n替代所\r\n需的大\r\n量SKU值)")
        self.ipLineEdit = QLineEdit()
        self.usrLineEdit = QLineEdit()
        self.pwdLineEdit = QLineEdit()
        self.portLineEdit = QLineEdit()
        self.databaseLineEdit = QLineEdit()
        self.excelLineEdit = QLineEdit()
        self.targetexcelLineEdit = QLineEdit()
        self.sqlTextEdit = QTextEdit()
        self.pwdLineEdit.setEchoMode(QLineEdit.Password)

        self.lb1 = QLabel(self)
        self.lb1.move(40, 10)
        self.lb1.setText('提示：sql类似 select * from table where goods_sn in (####)  SKU将会被excel 全部替换 \r\n       目标一定是csv格式 目标列一定是大写SKU三个字母')
        self.lb1.setStyleSheet("color:red")
        self.lb1.adjustSize()

        gridLayout = QGridLayout()
        gridLayout.addWidget(ip, 0, 0, 1, 1)
        gridLayout.addWidget(usr, 1, 0, 1, 1)
        gridLayout.addWidget(pwd, 2, 0, 1, 1)
        gridLayout.addWidget(port, 3, 0, 1, 1)
        gridLayout.addWidget(database, 4, 0, 1, 1)
        gridLayout.addWidget(excel, 5, 0, 1, 1)
        gridLayout.addWidget(targetexcel, 6, 0, 1, 1)
        gridLayout.addWidget(sql, 7, 0, 1, 1)
        gridLayout.addWidget(self.ipLineEdit, 0, 1, 1, 3)
        gridLayout.addWidget(self.usrLineEdit, 1, 1, 1, 3)
        gridLayout.addWidget(self.pwdLineEdit, 2, 1, 1, 3)
        gridLayout.addWidget(self.portLineEdit, 3, 1, 1, 3)
        gridLayout.addWidget(self.databaseLineEdit, 4, 1, 1, 3)
        gridLayout.addWidget(self.excelLineEdit, 5, 1, 1, 3)
        gridLayout.addWidget(self.targetexcelLineEdit, 6, 1, 1, 3)
        gridLayout.addWidget(self.sqlTextEdit, 7, 1, 1, 3)

        okBtn = QPushButton("导出数据")
        cancelBtn = QPushButton("退出")
        btnLayout = QHBoxLayout()

        btnLayout.setSpacing(60)
        btnLayout.addWidget(okBtn)
        btnLayout.addWidget(cancelBtn)

        dlgLayout = QVBoxLayout()
        dlgLayout.setContentsMargins(40, 40, 40, 40)
        dlgLayout.addLayout(gridLayout)
        dlgLayout.addStretch(40)
        dlgLayout.addLayout(btnLayout)

        self.setLayout(dlgLayout)
        okBtn.clicked.connect(self.accept)
        cancelBtn.clicked.connect(self.reject)
        self.setWindowTitle("专用数据Excel导出    BY: curry (专为大量excel的SKU而生)")
        self.resize(800, 600)

    def accept(self):
        if self.ipLineEdit.text().strip() == '' or self.usrLineEdit.text().strip() == '' \
            or self.pwdLineEdit.text().strip() =='' or self.portLineEdit.text().strip() == '' \
            or self.databaseLineEdit.text().strip() == '' or self.sqlTextEdit.toPlainText().strip() == ''\
            or self.excelLineEdit.text().strip() == '' or self.targetexcelLineEdit.text().strip() == '':
            QMessageBox.warning(self,
                    "提示",
                    "请完善信息！",
                    QMessageBox.Yes)
            self.ipLineEdit.setFocus()
        else:
            try:
                sku = pd.read_csv(self.excelLineEdit.text().strip(), encoding='gbk', dtype={'SKU':str}).dropna(axis=0)
                where = '","'.join(sku['SKU'])
                dbobj= pymysql.connect(
                    host=self.ipLineEdit.text().strip(),
                    database=self.databaseLineEdit.text().strip(),
                    user=self.usrLineEdit.text().strip(),
                    password=self.pwdLineEdit.text().strip(),
                    port=int(self.portLineEdit.text().strip()),
                    charset='utf8'
                    )
                selectsql = self.sqlTextEdit.toPlainText().strip().replace('####', '"'+where+'"')
                data = pd.read_sql(selectsql, dbobj)
                excel = pd.ExcelWriter(self.targetexcelLineEdit.text().strip())
                data.to_excel(excel)
                excel.save()
                QMessageBox.warning(self, "提示", "导出完成", QMessageBox.Yes)
                dbobj.close()
            except Exception as e:
                QMessageBox.warning(self, "提示", str(e), QMessageBox.Yes)
                self.ipLineEdit.setFocus()

app = QApplication(sys.argv)
dlg = LoginDlg()
dlg.show()
dlg.exec_()
app.exit()
