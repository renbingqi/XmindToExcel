# -*- ecoding: utf-8 -*-
# @ModuleName: main
# @Author: Rex
# @Time: 2023/1/15 11:04 下午
from core.handle_excel import HandleExcel
from core.handle_xmind import HandleXmind
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QFileDialog, QLineEdit, QMessageBox,QTabWidget
import sys
from PyQt5.QtGui import QIcon
import os


class Ui_MainWindow(QMainWindow, QTabWidget):
    file_path = None

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.resize(500, 300)
        self.label1 = QLabel(
            '<html><head/><body><p><span style=" color:#ff0000;">测试用例Xmind思维导图转Excel工具</span></p></body></html>）',
            self)
        self.label1.setGeometry(20, 25, 350, 16)
        # 选择xmind文件
        self.cblabel = QLabel("请选择xmind文件:", self)
        self.cblabel.setGeometry(20, 70, 179, 22)
        # 输入框
        self.pathlineEdit=QLineEdit(self)
        self.pathlineEdit .setGeometry(150, 70, 179, 22)
        # 选择xmind文件按钮
        self.selectButton=QPushButton("打开",self)
        self.selectButton.setGeometry(350, 65, 70, 30)
        self.selectButton.clicked.connect(self.select_file)
        #启动按钮
        self.startButton = QPushButton("开始转换", self)
        self.startButton.setGeometry(180, 140, 90, 40)
        self.startButton.clicked.connect(self.start)
        # 创建一个消息盒子，用于后续提示状态
        self.messagebox = QMessageBox()
        self.show()


    def select_file(self):
        dname = QFileDialog.getOpenFileName(self, 'Open file', 'C:\\Users\\test\\Desktop\\数据')
        self.file_path = dname[0]
        self.pathlineEdit.setText(self.file_path)

    def start(self):

        if  not self.file_path.endswith("xmind"):
            self.messagebox.information(self, "Message", "请选择正确的Xmind文件", QMessageBox.Ok)
        else:
            xmindHandler = HandleXmind(self.file_path)
            current_dir=os.path.dirname(self.file_path)
            xmindHandler.handle_xmind()
            fileName = xmindHandler.sheetName
            filePath = f"{current_dir}/{fileName}"
            maxModule = xmindHandler.maxModule
            data_list = xmindHandler.case_list
            excelHandler = HandleExcel(fileName, filePath)
            excelHandler.generate_title(maxModule)
            status=excelHandler.write_data(data_list)
            if status:
                self.messagebox.information(self, "Message", "文件转换完成", QMessageBox.Ok)
            else:
                self.messagebox.information(self, "Message", "文件转换失败，请检查xmind格式是否符合标准", QMessageBox.Ok)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("Pictures/icon.jpg"))
    ex = Ui_MainWindow()
    sys.exit(app.exec_())
