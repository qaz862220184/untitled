from PyQt5 import QtWidgets,QtCore
from PyQt5.QtWidgets import QTableWidget
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from case import Ui_Dialog
import sys,time
import xlrd
import asyncio,time,threading
import user,apply,register
import get_stu_edu,readExcel
class EduCaseQt(QTableWidget,QtWidgets.QWidget,Ui_Dialog):
    show_signal = pyqtSignal()
    def __init__(self):
        super(EduCaseQt,self).__init__()

        self.initUI()



    def initUI(self):
        self.setupUi(self)
        new = Ui_Dialog()
        new.setupUi(widget)
        self.pushButton_2.clicked.connect(self.click_success2)
        #self.table_widget()





class UpdateData(QThread):
    """更新数据类"""
    update_date = pyqtSignal(list)  # pyqt5 支持python3的str，没有Qstring

    def run(self):

        while True:
            data = readExcel.read_excel()
            self.update_date.emit(data)
            time.sleep(3)


class UpdateData2(QThread):
    """更新数据类"""
    update_date = pyqtSignal(list)  # pyqt5 支持python3的str，没有Qstring

    def run(self):

        while True:
            try:
                now = time.strftime("%d_%m_%Y")
                result_path = "./qt_%s.xlsx" % now
                workbook = xlrd.open_workbook(result_path)
                sheet = workbook.sheet_by_index(0)
                row_num = sheet.nrows
                col_num = sheet.ncols

                data1 = []
                data2 = []
                for i in range(1, row_num):
                    data = []
                    for j in range(col_num):
                        data.append(sheet.cell(i, j).value)
                    data1.append(data)
                data2.append(row_num)
                data2.append(col_num)
                data2.append(data1)
            except FileNotFoundError:
                data2=[1,2,3]
            self.update_date.emit(data2)
            time.sleep(3)







if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    widget = QtWidgets.QWidget()
    new = Ui_Dialog()
    new.setupUi(widget)



    update_data_thread = UpdateData()
    update_data_thread.update_date.connect(new.update_item_data)  # 链接信号
    update_data_thread.start()

    update_data_thread2 = UpdateData2()
    update_data_thread2.update_date.connect(new.update_item_data2)  # 链接信号
    update_data_thread2.start()

    # new_loop = asyncio.new_event_loop()  # 在当前线程下创建时间循环，（未启用）
    # t = threading.Thread(target=start_loop, args=(new_loop,))  # 开启新的线程去启动事件循环
    # t.start()
    # asyncio.run_coroutine_threadsafe(new.click_success2,new_loop)
    #asyncio.set_event_loop(asyncio.new_event_loop())

    new.pushButton_2.clicked.connect(new.click_success2)
    new.pushButton_3.clicked.connect(new.click_success3)

    widget.show()
    sys.exit(app.exec())


