from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import QEvent, QThread, Qt, QEventLoop, QTimer, pyqtSignal
import os
import glob
import time
import pandas as pd
from logic import lg


class Ui_Form_Add(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(301, 213)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(20, 10, 261, 151))
        self.groupBox.setObjectName("groupBox")
        self.listExcelAdd = QtWidgets.QListWidget(self.groupBox)
        self.listExcelAdd.setGeometry(QtCore.QRect(10, 20, 241, 121))
        self.listExcelAdd.setObjectName("listExcelAdd")
        self.addNow = QtWidgets.QPushButton(Form)
        self.addNow.setGeometry(QtCore.QRect(70, 170, 161, 31))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.addNow.setFont(font)
        self.addNow.setObjectName("addNow")
        self.addNow.setStyleSheet('''
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

        file_list = glob.glob("*.xlsx")
        file_list.sort(key=os.path.getmtime, reverse=True)  # sorted by modified date
        for each_file in file_list:
            if not ("To Do _ " in each_file or "Thống Kê _ " in each_file):
                self.listExcelAdd.addItem(each_file)

        self.addNow.clicked.connect(self.evt_btnAdd_clicked)

        # self.msgBox = QtWidgets.QMessageBox(self)
        # self.msgBox.setWindowTitle("Title")
        # self.msgBox.setText("Start")
        # self.msgBox.show()

    def evt_btnAdd_clicked(self):

        excel_name = self.listExcelAdd.currentItem().text()

        self.worker = WorkerThreadAdd(excel_name)
        self.worker.start()
        self.worker.add_finish.connect(self.evt_worker_finished)
        self.worker.add_progress.connect(self.evt_worker_update)
        self.worker.add_check_ghi_de.connect(self.evt_check_ghi_de)
        self.worker.add_error.connect(self.evt_worker_error)

    def evt_worker_error(self, error):
        lg.alert("Lỗi", error)

    def evt_check_ghi_de(self, titleTaskList):
        if (titleTaskList!=""):
            msg = QtWidgets.QMessageBox()
            result = QtWidgets.QMessageBox.question(msg, "Bạn có muốn ghi đè không?", "Việc này sẽ làm XÓA HẾT tasks list \"" + titleTaskList +
                                                    "\" cũ và thay thế cái mới", QtWidgets.QMessageBox.StandardButton.Ok, QtWidgets.QMessageBox.StandardButton.Cancel)
            if (result == QtWidgets.QMessageBox.StandardButton.Ok):
                self.worker.receive_check_ghi_de = 1
            else:
                self.worker.receive_check_ghi_de = 0

    def evt_worker_finished(self):
        lg.alert("Success!", "Add thành công hết rồi nè <3")

    def evt_worker_update(self, log):
        self.progressLb.setText(log)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Thêm việc cần làm"))
        self.groupBox.setTitle(_translate("Form", "Danh sách file excel"))
        self.addNow.setText(_translate("Form", "ADD TO GOOGLE TASKS"))


class WorkerThreadAdd(QThread):
    add_progress = pyqtSignal(str)
    add_error = pyqtSignal(str)
    add_check_ghi_de = pyqtSignal(str)
    receive_check_ghi_de = -1  # Trạng thái chờ nhận tín hiệu Check Ghi đè
    add_finish = pyqtSignal(bool)

    def __init__(self, excel_name):
        QThread.__init__(self)
        self.excel_name = excel_name

    def run(self):  # Hàm Update

        # Bước 1: Fill time và export 2 cột taskTitleColumn, taskDueColumn
        # excel_name = 'Toán thầy Chí test.xlsx'
        excel_name_todo = "To Do _ " + \
            self.excel_name[:-5]+".xlsx"   # Tên file excel to do

        df = pd.read_excel(self.excel_name, header=None,
                           names=['taskTitle', 'taskDue'])

        fill = lg.fill_time(df, excel_name_todo)
        if isinstance(fill, str):   # Gửi lỗi
            self.add_error.emit(fill)
            return

        taskTitleColumn = fill["title"]
        taskDueColumn = fill["due"]

        # Bước 2: Tạo task list theo tên Excel file, ghi đè nếu trùng tên

        titleTaskList = self.excel_name[:-5]

        listTaskLists = lg.get_lists_task()
        listTaskLists_name = [d["title"] for d in listTaskLists]

        # Check ghi đè Tasks List, Nếu người dùng Accept ghi đề thì xóa Task List cũ -> Tạo Task List mới

        if not titleTaskList in listTaskLists_name:    # Nếu chưa tồn tại thì tạo project mới
            mainTaskListId = lg.add_lists_tasks(titleTaskList).get('id')

        else:    # Nếu tồn tại
            self.add_check_ghi_de.emit(titleTaskList)
            while (self.receive_check_ghi_de == -1):
                # Chờ nhận tín hiệu trả về của Dialog Ghi đè (Vương thông minh vãi :))
                time.sleep(0.3)

            if (self.receive_check_ghi_de == 1):
                print("ghi đè")
                mainTaskListId = listTaskLists[listTaskLists_name.index(
                    titleTaskList)].get('id')

                lg.del_list_task(mainTaskListId)

                mainTaskListId = lg.add_lists_tasks(titleTaskList).get('id')
            else:
                print("Hủy bỏ")
                return

        # Bước 3: Bắt đầu Add hàng loạt to Google Tasks {title, due}
        count = 1
        totalCount = len(taskTitleColumn)
        for taskTitle, taskDue in zip(reversed(taskTitleColumn), reversed(taskDueColumn)):
            while True:  # Tránh lỗi Quota Exceeded
                try:
                    lg.add_task(mainTaskListId=mainTaskListId,
                                 taskTitle=taskTitle, taskDue=taskDue)
                    log1 = "{} / {}".format(count, totalCount)
                    # log = "{}\n Success updated! {}\n {}\n".format(log1,taskTitle,taskDue.strftime("%d/%m/%Y"))
                    print(log1)
                    self.add_progress.emit(log1)
                except Exception as e:
                    print(e)
                    # log = ("Google \"quá tải\" rồi... Vui lòng đợi 18s nhé <3")
                    log = "Đợi xíu..."
                    self.add_progress.emit(log)
                    time.sleep(18)
                    continue
                break
            count += 1

        self.add_finish.emit(True)
        # Gửi tín hiệu Add thành công


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form_Add()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())
