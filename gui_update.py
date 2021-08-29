from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import QEvent, QThread, Qt, QEventLoop, QTimer, pyqtSignal
import os
import glob
import time
import pandas as pd
from pandas.core.series import Series
from logic import lg

class Ui_Form_Update(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(301, 213)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(20, 10, 261, 151))
        self.groupBox.setObjectName("groupBox")
        self.listExcelUpdate = QtWidgets.QListWidget(self.groupBox)
        self.listExcelUpdate.setGeometry(QtCore.QRect(10, 20, 241, 121))
        self.listExcelUpdate.setObjectName("listExcelUpdate")
        self.updateNow = QtWidgets.QPushButton(Form)
        self.updateNow.setGeometry(QtCore.QRect(80, 170, 141, 31))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.updateNow.setFont(font)
        self.updateNow.setObjectName("updateNow")
        self.updateNow.setStyleSheet('''
            QPushButton {
                font-weight: bold;
            }
        ''')

        self.progressLb = QtWidgets.QLabel(Form)
        self.progressLb.setGeometry(QtCore.QRect(243, 191, 47, 13))
        self.progressLb.setText("")
        self.progressLb.setObjectName("progressLb")
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
#________________________________________Ở trên là Generate By Qt Designer_____________________________________________________

        file_list = glob.glob("*.xlsx")
        file_list.sort(key=os.path.getmtime, reverse=True)  # sorted by modified date
        for each_file in file_list:
            if "Thống Kê _ " in each_file:
                self.listExcelUpdate.addItem(each_file)
                
        self.updateNow.clicked.connect(self.evt_btnUpdate_clicked)

        # self.msgBox = QtWidgets.QMessageBox(self)
        # self.msgBox.setWindowTitle("Title")
        # self.msgBox.setText("Start")
        # self.msgBox.show()

    def evt_btnUpdate_clicked(self):              
        
        excel_name = self.listExcelUpdate.currentItem().text()
        print(excel_name)
        # if (len(excel_name)==0):
        #     lg.alert("Chưa chọn file", "Chọn file rồi mới UPDATE được chứ! <3")
        #     return
                
        self.worker = WorkerThreadUpdate(excel_name)
        self.worker.start()
        self.worker.update_finish.connect(self.evt_worker_finished)
        self.worker.update_progress.connect(self.evt_worker_update)
        self.worker.update_error.connect(self.evt_worker_error)

    def evt_worker_error(self, error):
        lg.alert("Lỗi", error)

    def evt_worker_finished(self, check):
        if check:
            lg.alert("Updated thành công!", "Thành công rồi nè <3")
        else:
            lg.alert("Có lỗi xảy ra", "Bạn vui lòng thử lại sau hoặc liên hệ hỗ trợ nhé")

    def evt_worker_update(self, log):
        self.progressLb.setText(log)
        # self.msgBox.setText(log)
        # loop = QEventLoop()
        # QTimer.singleShot(1000, loop.quit)
        # loop.exec()

    def retranslateUi(self, Form):  # ________ Đoạn này cũng Generate by Qt Designer
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Cập nhật việc cần làm"))
        self.groupBox.setTitle(_translate("Form", "Danh sách file excel"))
        self.updateNow.setText(_translate("Form", "UPDATE NOW"))

  
# ___________________________ Thread UPDATE _________________________
class WorkerThreadUpdate(QThread):
    update_progress = pyqtSignal(str)
    update_error = pyqtSignal(str)
    update_finish = pyqtSignal(bool)

    def __init__(self, excel_name):
        QThread.__init__(self)
        self.excel_name = excel_name

    def run(self):  # Hàm Update
        # excel_name = 'Thống Kê _ Hack não Ielts B.xlsx'
        excel_name_todo =  "To Do _ " + self.excel_name[self.excel_name.index("_ ")+2:]

        df = pd.read_excel(self.excel_name)
        mainTaskListId = df.columns[0]
        df.columns=['id','taskTitle', 'taskDue','total']

        fill = lg.fill_time(df,excel_name_todo, isUpdate=True)
        if isinstance(fill, str):   # Gửi lỗi
            self.update_error.emit(fill)
            return

        taskTitleColumn = fill["title"]
        taskDueColumn = fill["due"]
        taskIdColumn = list(Series(df['id']))

        count=1
        totalCount = len(taskIdColumn)
        # Bắt đầu Update hàng loạt to Google Tasks {title, due}
        for taskId, taskTitle, taskDue in zip(reversed(taskIdColumn), reversed(taskTitleColumn), reversed(taskDueColumn)):
            # while True:  # Tránh lỗi Quota Exceeded
            try:
                response = lg.get_task(mainTaskListId,taskId)
                response['title']=taskTitle
                response['due']=lg.convert_VietNam_Date_Object_to_RFC(taskDue)
                lg.update_tasks(mainTaskListId, taskId, response)

                log1 = "{} / {}".format(count,totalCount)
                log = "{}\n Success updated! {}\n {}\n".format(log1,taskTitle,taskDue.strftime("%d/%m/%Y"))
                print(log1)
                self.update_progress.emit(log)
            except Exception as e: 
                print(e)
                self.update_finish.emit(False)
                return
                # continue
            count += 1
        # Update Thành công
        
        self.update_finish.emit(True)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form_Update()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())

