import pandas as pd
from pandas.core.series import Series
from dateutil.parser import parse
import pytz
from PyQt6 import QtWidgets
from Google import Create_Service
import os

# ______________________ Initiate Google Tasks ______________________________
assetsFolder = os.path.join(os.getcwd(), 'assets')

CLIENT_SECRET_FILE = os.path.join(assetsFolder, 'login.json')
API_NAME = 'tasks'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/tasks']
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# Hàm định dạng thời gian
def strfdelta(tdelta, fmt):
    d = {"days": tdelta.days}
    d["hours"], rem = divmod(tdelta.seconds, 3600)
    d["minutes"], d["seconds"] = divmod(rem, 60)
    return fmt.format(**d)

class lg:
    def __init__(self):
        print("in init")

    def alert(title,content,moreContent=""):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(content)
        msg.setInformativeText(moreContent)
        msg.exec()
        return msg

    def convert_VietNam_Date_Object_to_RFC(vietnamDate):
        return vietnamDate.isoformat() + 'Z'

        # utc_dt = vietnamDate.astimezone(pytz.utc) Trường hợp cần chuyển múi giờ -7
        # rfc_dt = utc_dt.isoformat("T") + 'Z'
        # return rfc_dt.replace('+00:00','')

    def convert_RFC_to_VietNam_Date_Object(rfcDate):
        VN_TZ = pytz.timezone('Asia/Ho_Chi_Minh')
        date_time_vietnam = parse(rfcDate).astimezone(VN_TZ)
        return date_time_vietnam

    # Hàm điền thời gian theo cấp số cộng
    def fill_time(df, excel_name, isUpdate=False):
        # Trích xuất cột Title và cột Due thành 2 List
        taskTitleColumn = list(Series(df['taskTitle']))
        taskDueColumn = list(Series(df['taskDue']))

        # Kiếm tra BEGIN và START thử người dùng có nhập chưa
        if (pd.isnull(taskDueColumn[0]) or pd.isnull(taskDueColumn[-1])):
            return ("Nhập thiếu ngày \"bắt đầu / kết thúc\" rùi kìa <3")

        # Convert Date Time in Excel to native Python datetime object
        
        for i, dueItem in enumerate(taskDueColumn):
            try:
                if isinstance(dueItem, str):   # Nếu cột là str (trường hợp UPDATE)
                    taskDueColumn[i] = pd.to_datetime(dueItem, dayfirst=True).to_pydatetime()
                else:  # Nếu cột thuộc kiểu Timestamp
                    taskDueColumn[i] = pd.Timestamp(dueItem).to_pydatetime()
            except:
                return ("File excel lỗi định dạng ngày tháng rồi, sửa lại đi <3")

        # index BEGIN, END datetime in excel (lấy vị trí để lặp từng khúc)
        flag = []
        for index, eachDate in enumerate(taskDueColumn):
            if (not pd.isnull(eachDate)):
                flag.append(index)

        listLog = []
        for j in range(0, len(flag)-1, 1):  # Lặp từng khúc để rải đều thời gian
            begin = flag[j]
            end = flag[j+1]
            if (begin+1 == end):  # Nếu begin và end đứng cạnh nhau => khỏi cần fill
                continue
            
            # IF ERROR ngày sau nhỏ hơn ngày trước
            if ((taskDueColumn[end] - taskDueColumn[begin]).days < 0):
                vtBegin = begin + 1
                vtEnd = end + 1
                if isUpdate:    # Vì có thêm Header nên phải cộng thêm 1
                    vtBegin += 1
                    vtEnd += 1
                return ("Nhập sai ngày tháng rồi kìa <3! \nNgày kết thúc NHỎ HƠN ngày bắt đầu ở hàng {} và hàng {}".format(vtBegin, vtEnd))

            # d = (Un - U1) / (n-1)    (Kiến thức toán cấp số cộng)
            stepDate = (taskDueColumn[end] - taskDueColumn[begin]) / (end-begin)

            stepDatePrint = strfdelta(stepDate, "\"{days} ngày {hours} tiếng\"")
            totalDatePrint = strfdelta(
                taskDueColumn[end]-taskDueColumn[begin], "\"{days} NGÀY\"")
            listLog.append("Từ task {} ĐẾN task {} \n Bạn có khoảng {} để hoàn thành mỗi task nhé <3. Tổng thời gian khoảng: {}".format(
                [begin+2], [end+2], stepDatePrint, totalDatePrint))

            for i in range(begin+1, end):   # Bắt đầu rải thời gian theo cấp số cộng
                listLog.append("")
                taskDueColumn[i] = taskDueColumn[i-1]+stepDate
            listLog.append("")

        # Write to Excel file
        try:
            writer = pd.ExcelWriter(excel_name, engine='xlsxwriter')
        except:
            return ("Vui lòng tắt cửa số "+excel_name)
            # time.sleep(5)
            # continue

        df = pd.DataFrame(list(
            zip(taskTitleColumn, [x.strftime("%d/%m/%Y") for x in taskDueColumn], listLog)))
        df.columns = ['Task', 'Hạn chót', 'Ước lượng thời gian để hoàn thành']

        df.to_excel(writer, sheet_name='Nhiệm vụ', index=False)

        # Auto-adjust columns width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Nhiệm vụ'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Nhiệm vụ'].set_column(
            0, 0, 60)  # Set column Title width 60px

        writer.save()
        writer.close()

        return {"title": taskTitleColumn, "due": taskDueColumn}


    # Hàm get tasks list (project)
    def get_lists_task():
        response = service.tasklists().list().execute()
        listItems = response.get('items')
        nextPageToken = response.get('nextPageToken')
        while nextPageToken:
            response = service.tasklists().list(
                maxResults=30,
                pageToken=nextPageToken
            ).execute()
            listItems.extend(response.get('items'))
            nextPageToken = response.get('nextPageToken')
        return listItems


    # Hàm add tasks list (project), trả về object tasklist
    def add_lists_tasks(titleTaskList):
        taskList = service.tasklists().insert(
            body={'title': titleTaskList}
        ).execute()
        return taskList

    # Hàm xóa tasks list
    def del_list_task(mainTaskListId):
        service.tasklists().delete(tasklist=mainTaskListId).execute()

    # Hàm lấy các Task trong mainTaskListId
    def get_tasks(mainTaskListId, complete=True, hidden=True):
        response = service.tasks().list(
            tasklist=mainTaskListId,
            showCompleted=complete,
            showHidden=hidden,
        ).execute()
        listItems = response.get('items')
        nextPageToken = response.get('nextPageToken')

        while nextPageToken:
            response = service.tasks().list(
                tasklist=mainTaskListId,
                showCompleted=complete,
                showHidden=hidden,
                pageToken=nextPageToken
            ).execute()
            listItems.extend(response.get('items'))
            nextPageToken = response.get('nextPageToken')

        listItems_sortPosition = sorted(listItems, key=lambda k: k['position'])
        return listItems_sortPosition

    # Hàm get 1 task
    def get_task(tasklistID, taskID):
        response = service.tasks().get(
            tasklist = tasklistID,
            task = taskID,
        ).execute()
        return response

    # Hàm ADD Task
    def add_task(mainTaskListId, taskTitle, taskDue):
        response = service.tasks().insert(
            tasklist=mainTaskListId,
            body={
                'title': taskTitle,
                'due': taskDue.isoformat() + 'Z'
            }
        ).execute()
        return response

    # Hàm UPDATE Task
    def update_task(mainTaskListId, taskID, bodyUpdate):
        response = service.tasks().update(
            tasklist=mainTaskListId,
            task=taskID,
            body=bodyUpdate
        ).execute()
        return response

    # Hàm DELETE Tasks
    def delete_task(mainTaskListId, taskId):
        service.tasks().delete(tasklist=mainTaskListId, task=taskId).execute()

