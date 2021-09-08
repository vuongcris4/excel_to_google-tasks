# pyinstaller gui_main.py --onefile --noconsole --icon="assets\main.ico"

import sys
import time
from PyQt6 import QtWidgets
from PyQt6.QtGui import QGuiApplication, QIcon
from PyQt6.QtWidgets import QApplication, QDial, QLCDNumber, QMenu, QMessageBox, QWidget, QHBoxLayout, QVBoxLayout
from PyQt6.QtCore import QEvent, QThread, Qt, QEventLoop, QTimer, pyqtSignal
from PyQt6 import uic
from gui_add import Ui_Form_Add
from gui_update import Ui_Form_Update
from logic import lg
import pandas as pd
import os
import webbrowser

assetsFolder = os.path.join(os.getcwd(), 'assets')
with open(os.path.join(assetsFolder, 'urlHuongDan.txt'), encoding='utf-8') as f:
    urlHuongDan = f.read()
    f.close()

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi(os.path.join(assetsFolder, 'gui_main.ui'), self)
        self.setWindowIcon(QIcon(os.path.join(assetsFolder, "main.ico")))

        self.DanhSachViecCanLam()
        
        self.reloadBtn.clicked.connect(self.DanhSachViecCanLam)
        self.addTasks.clicked.connect(self.openWindowAdd)  # Bấm vào thì mở cửa sổ ADD
        self.updateTasks.clicked.connect(self.openWindowUpdate)  # Bấm vào thì mở cửa sổ UPDATE
        self.hdsdBtn.clicked.connect(self.hdsd)
        self.logOutBtn.clicked.connect(self.logOut)
    
    def hdsd(self):
        webbrowser.open(urlHuongDan)  

    def logOut(self):
        os.remove(os.path.join(assetsFolder, "file_chua_thong_tin_dang_nhap.pickle"))
        sys.exit()

    def DanhSachViecCanLam(self):
        self.listTasks.clear()
        for listTaskName in self.GET_listTaskLists_name():    # Khởi tạo List Tasks Name
            self.listTasks.addItem(listTaskName)
        self.listTasks.installEventFilter(self)   # Tạo Context Menu "Lấy file thống kê"

    def GET_listTaskLists_name(self):
        listTaskLists = lg.get_lists_task()
        listTaskLists_name = [d["title"] for d in listTaskLists]
        return listTaskLists_name

    def eventFilter(self, source, event):   # Context Menu GET File Thống Kê
        if event.type() == QEvent.Type.ContextMenu and source is self.listTasks:
            menu = QMenu()
            menu.addAction("Lấy file thống kê")

            if menu.exec(event.globalPos()):
                item = source.itemAt(event.pos())
                selected = item.text()
                print(selected)

                self.worker = WorkerThread_GET_file_Thong_Ke(selected)
                self.worker.start()

                self.worker.get_finish.connect(self.evt_worker_finished)
                self.worker.get_error.connect(self.evt_worker_error)

            return True
        return super().eventFilter(source, event)   # Context Menu

    def evt_worker_error(self, error):
        lg.alert("Lỗi", error)

    def evt_worker_finished(self,excelNameUpdate):  # Khi thành công    GET File Thống Kê
        lg.alert("Thành công","Đã xuất file \""+excelNameUpdate+"\" thành công! <3")

    def openWindowAdd(self):
        self.window  = QtWidgets.QMainWindow()
        self.ui = Ui_Form_Add()
        self.ui.setupUi(self.window)
        self.window.show()  # Thêm tính năng Reload

    def openWindowUpdate(self):
        self.window  = QtWidgets.QMainWindow()
        self.ui = Ui_Form_Update()
        self.ui.setupUi(self.window)
        self.window.show()


class WorkerThread_GET_file_Thong_Ke(QThread):
    get_error = pyqtSignal(str)
    get_finish = pyqtSignal(str)

    def __init__(self, selectedTaskList):
        QThread.__init__(self)
        self.selectedTaskList = selectedTaskList
          
    def run(self):
        listTaskLists = lg.get_lists_task()
        listTaskLists_name = [d["title"] for d in listTaskLists]

        # Bước này in listTaskLists_name ra GUI để người dùng chọn ___                       CẦN LÀM

        # selectedTaskList = 'Toán thầy Chí test'  # Giả sử chọn
        excelNameUpdate = "Thống Kê _ " + self.selectedTaskList+".xlsx"

        mainTaskListId = listTaskLists[listTaskLists_name.index(
            self.selectedTaskList)].get('id')

        # Lấy các tasks trong project mainTaskListId
        listTaskList_main = lg.get_tasks(mainTaskListId)

        # "needsAction" or "completed"
        listTaskList_main_completed = [
            x for x in listTaskList_main if x['status'] == "completed"]
        listTaskList_main_uncompleted = [
            x for x in listTaskList_main if x['status'] == "needsAction"]

        #  Xử lí Sheets
        # _______________________-Xuất Excel THỐNG KÊ

        options = {}    # Tránh để lỗi khi gặp dấu bằng, bloabloa
        options['strings_to_formulas'] = False
        options['strings_to_urls'] = False

        # os.path.join(DATA_DIR, 'Data.xlsx')
        while True:
            try:
                writer = pd.ExcelWriter(excelNameUpdate, engine='xlsxwriter')
            except:
                self.get_error.emit("Vui lòng tắt cửa số \""+excelNameUpdate+"\" \n\nTắt đi để tui còn lưu file lại nữa nè <3")
                # print("Vui lòng tắt cửa số "+excelNameUpdate)
                return 
            break

        
        # _______________ Sheet Chưa hoàn thành

        df_uncompleted = pd.DataFrame(listTaskList_main_uncompleted)
        try:
            due = [lg.convert_RFC_to_VietNam_Date_Object(x['due']) for x in listTaskList_main_uncompleted]
        except:
            self.get_error.emit("Bắt buộc cái task đều phải có Due Date \n\nBạn vui lòng bổ sung Due Date cho đầy đủ nhé <3")
            return
            # print("Bắt buộc cái task đều phải có Due Date")
            # sys.exit(1)


        # Xử lí để chỉ còn lại Begin và End time task
        dueEmpty = [due[0].strftime("%d/%m/%Y")]
        for i in range(1, len(due)-1):
            if (due[i-1] <= due[i] and due[i] <= due[i+1]):
                dueEmpty.append("")
            else:
                dueEmpty.append(due[i].strftime("%d/%m/%Y"))
        dueEmpty.append(due[-1].strftime("%d/%m/%Y"))

        df_uncompleted['due'] = dueEmpty
        df_uncompleted = df_uncompleted[['id', 'title', 'due']]
        df_uncompleted.columns = [mainTaskListId, 'Task', 'Hạn chót']   # Gán ID Tasks List 
        df_uncompleted["Bạn còn phải thực hiện {} / {} tasks <3".format(len(listTaskList_main_uncompleted),len(listTaskList_main))]="" # In cột tổng: Còn phải thực hiện ../.. tasks
        # df_uncompleted = df_uncompleted.iloc[::-1]   # Reverse DataFrame


        df_uncompleted.to_excel(writer, sheet_name='Đang thực hiện', index=False)

        # Auto-adjust columns width sheet 'Đang thực hiện'
        for column in df_uncompleted:
            column_width = max(df_uncompleted[column].astype(
                str).map(len).max(), len(column))
            col_idx = df_uncompleted.columns.get_loc(column)
            writer.sheets['Đang thực hiện'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Đang thực hiện'].set_column(0, 0, 0)  # Hide column ID


        # _______________ Sheet ĐÃ hoàn thành
        if len(listTaskList_main_completed)!=0: # Nếu có task hoàn thành rồi thì mới tạo sheet đã hoàn thành
            df_completed = pd.DataFrame((listTaskList_main_completed))
            
            completed = [lg.convert_RFC_to_VietNam_Date_Object(x['completed']).strftime("%d/%m/%Y %H:%M")
                        for x in listTaskList_main_completed]    # Chuyển RFC Time về dạng time Việt Nam

            df_completed['completed'] = (completed)
            df_completed = df_completed[['title', 'completed']]
            df_completed.columns = [
                'Task', 'ĐÃ HOÀN THÀNH vào lúc']
            df_completed["Bạn ĐÃ HOÀN THÀNH {} / {} tasks <3".format(len(listTaskList_main_completed),len(listTaskList_main))]="" # In cột tổng: Đã hoàn thành ../.. tasks

            df_completed.to_excel(writer, sheet_name='Đã hoàn thành', index=False)
            
            # Auto-adjust columns width sheet 'Đã hoàn thành'
            for column in df_completed:
                column_width = max(df_completed[column].astype(
                    str).map(len).max(), len(column))
                col_idx = df_completed.columns.get_loc(column)
                writer.sheets['Đã hoàn thành'].set_column(col_idx, col_idx, column_width)

            # Format center align cột due và cột completed
            fmt = writer.book.add_format({'align': 'center', 'valign': 'vcenter'})    #, 'text_wrap': True
            writer.sheets['Đã hoàn thành'].set_column(1,1,25,fmt)       # center Align

        writer.save()
        writer.close()

        self.get_finish.emit(excelNameUpdate)
        # Gửi tín hiệu Lấy file Excel thành công    (gửi tên file)
        return

if __name__ == '__main__':
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    myApp = MyApp()
    myApp.show()
    try:
        sys.exit(app.exec())
    except SystemExit:
        print('Closing Window...')
