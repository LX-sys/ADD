# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ADD.ui'
#
# Created by: PyQt5 UI code generator 5.15.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.
import sys
import time

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QDate, QCalendar,Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QHeaderView, QAbstractItemView, QFileDialog, \
    QMessageBox, QTableWidget
import pprint

from ChildAddDyNumber import CdyNumber
from myExcel import MyExcel
from ADDtype import TYpe
import QSS

class MainUI(QMainWindow):

    def __init__(self):
        super(MainUI, self).__init__()
        # 下拉框的文本个数
        self._comboBoxIndex = 0
        # 历史索引
        self._historyIndex = 0
        # 下拉框的文本
        self._comboBoxText = []
        # tab的当前值
        self._tabValueing = None
        self.choose ={"name":"","r":1}
        self.setupUi()

    def setupUi(self):
        self.setObjectName("main")
        self.resize(1023, 776)
        # 窗口居中
        screen = QtWidgets.QDesktopWidget().screenGeometry()  # 获取屏幕分辨率
        size = self.geometry()  # 获取窗口尺寸
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)  # 利用move函数窗口居中

        # font = QtGui.QFont("黑体")
        # font.setBold(False)
        # font.setWeight(50)
        # self.setFont(font)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_8 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.top_widget = QtWidgets.QWidget(self.centralwidget)
        self.top_widget.setObjectName("top_widget")
        self.gridLayout_6 = QtWidgets.QGridLayout(self.top_widget)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.serialNum_label = QtWidgets.QLabel(self.top_widget)
        self.serialNum_label.setMinimumSize(QtCore.QSize(51, 24))
        self.serialNum_label.setMaximumSize(QtCore.QSize(71, 24))
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.serialNum_label.setFont(font)
        self.serialNum_label.setObjectName("serialNum_label")
        self.horizontalLayout.addWidget(self.serialNum_label)
        self.serialNum_comboBox = QtWidgets.QComboBox(self.top_widget)
        self.serialNum_comboBox.setMinimumSize(QtCore.QSize(148, 32))
        font = QtGui.QFont("黑体")
        font.setFamily(".Keyboard")
        font.setBold(False)
        font.setWeight(50)
        self.serialNum_comboBox.setFont(font)
        self.serialNum_comboBox.setObjectName("serialNum_comboBox")
        self.horizontalLayout.addWidget(self.serialNum_comboBox)
        self.serialNum_addPush = QtWidgets.QPushButton(self.top_widget)
        self.serialNum_addPush.setMinimumSize(QtCore.QSize(81, 32))
        self.serialNum_addPush.setMaximumSize(QtCore.QSize(81, 32))
        self.serialNum_addPush.setObjectName("serialNum_addPush")
        self.horizontalLayout.addWidget(self.serialNum_addPush)
        self.gridLayout_6.addLayout(self.horizontalLayout, 0, 0, 1, 1)
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.excelName_label = QtWidgets.QLabel(self.top_widget)
        self.excelName_label.setMinimumSize(QtCore.QSize(71, 16))
        self.excelName_label.setMaximumSize(QtCore.QSize(16, 16777215))
        self.excelName_label.setObjectName("excelName_label")
        self.gridLayout_3.addWidget(self.excelName_label, 0, 0, 1, 1)
        self.excelShowNamelabel = QtWidgets.QLabel(self.top_widget)
        self.excelShowNamelabel.setMinimumSize(QtCore.QSize(161, 16))
        self.excelShowNamelabel.setObjectName("excelShowNamelabel")
        self.gridLayout_3.addWidget(self.excelShowNamelabel, 0, 1, 1, 1)
        self.excelPos_label = QtWidgets.QLabel(self.top_widget)
        self.excelPos_label.setMinimumSize(QtCore.QSize(61, 16))
        self.excelPos_label.setMaximumSize(QtCore.QSize(61, 16))
        self.excelPos_label.setObjectName("excelPos_label")
        self.gridLayout_3.addWidget(self.excelPos_label, 1, 0, 1, 1)
        self.cecelShowPos_label = QtWidgets.QLabel(self.top_widget)
        self.cecelShowPos_label.setMinimumSize(QtCore.QSize(171, 16))
        self.cecelShowPos_label.setObjectName("cecelShowPos_label")
        self.gridLayout_3.addWidget(self.cecelShowPos_label, 1, 1, 1, 1)
        self.gridLayout_6.addLayout(self.gridLayout_3, 0, 1, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.top_widget)
        self.pushButton.setMinimumSize(QtCore.QSize(41, 41))
        self.pushButton.setMaximumSize(QtCore.QSize(41, 41))
        self.pushButton.setObjectName("pushButton")
        self.gridLayout_6.addWidget(self.pushButton, 0, 2, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(169, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_6.addItem(spacerItem, 0, 3, 1, 1)
        self.gridLayout_8.addWidget(self.top_widget, 0, 0, 1, 1)
        self.splitter = QtWidgets.QSplitter(self.centralwidget)
        self.splitter.setOrientation(Qt.Vertical)
        self.splitter.setObjectName("splitter")
        self.middle_widget = QtWidgets.QWidget(self.splitter)
        self.middle_widget.setObjectName("middle_widget")
        self.gridLayout_10 = QtWidgets.QGridLayout(self.middle_widget)
        self.gridLayout_10.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.calendarWidget = QtWidgets.QCalendarWidget(self.middle_widget)
        self.calendarWidget.setObjectName("calendarWidget")
        self.gridLayout_10.addWidget(self.calendarWidget, 0, 0, 3, 1)
        self.init_pushButton = QtWidgets.QPushButton(self.middle_widget)
        self.init_pushButton.setMinimumSize(QtCore.QSize(90, 50))
        self.init_pushButton.setMaximumSize(QtCore.QSize(90, 50))
        self.init_pushButton.setObjectName("init_pushButton")
        self.gridLayout_10.addWidget(self.init_pushButton, 0, 1, 1, 1)
        self.groupBox = QtWidgets.QGroupBox(self.middle_widget)
        self.groupBox.setMinimumSize(QtCore.QSize(391, 261))
        self.groupBox.setObjectName("groupBox")
        self.gridLayout_7 = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout_7.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.scrollArea = QtWidgets.QScrollArea(self.groupBox)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 383, 234))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.gridLayout_12 = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout_12.setObjectName("gridLayout_12")
        self.gridLayout_11 = QtWidgets.QGridLayout()
        self.gridLayout_11.setObjectName("gridLayout_11")
        self.reply_pushButton_2 = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.reply_pushButton_2.setMinimumSize(QtCore.QSize(81, 50))
        self.reply_pushButton_2.setMaximumSize(QtCore.QSize(81, 50))
        font = QtGui.QFont("黑体")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.reply_pushButton_2.setFont(font)
        self.reply_pushButton_2.setObjectName("reply_pushButton_2")
        self.gridLayout_11.addWidget(self.reply_pushButton_2, 1, 1, 1, 1)
        self.reply_pushButton = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.reply_pushButton.setMinimumSize(QtCore.QSize(81, 50))
        self.reply_pushButton.setMaximumSize(QtCore.QSize(81, 50))
        font = QtGui.QFont("黑体")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.reply_pushButton.setFont(font)
        self.reply_pushButton.setObjectName("reply_pushButton")
        self.gridLayout_11.addWidget(self.reply_pushButton, 1, 0, 1, 1)
        self.hello_pushButton = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.hello_pushButton.setMinimumSize(QtCore.QSize(81, 50))
        self.hello_pushButton.setMaximumSize(QtCore.QSize(81, 50))
        font = QtGui.QFont("黑体")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.hello_pushButton.setFont(font)
        self.hello_pushButton.setObjectName("hello_pushButton")
        self.gridLayout_11.addWidget(self.hello_pushButton, 0, 0, 1, 1)
        self.hello_pushButton_2 = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.hello_pushButton_2.setMinimumSize(QtCore.QSize(81, 50))
        self.hello_pushButton_2.setMaximumSize(QtCore.QSize(81, 50))
        font = QtGui.QFont("黑体")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.hello_pushButton_2.setFont(font)
        self.hello_pushButton_2.setObjectName("hello_pushButton_2")
        self.gridLayout_11.addWidget(self.hello_pushButton_2, 0, 1, 1, 1)
        self.gridLayout_12.addLayout(self.gridLayout_11, 0, 0, 1, 1)
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.gridLayout_7.addWidget(self.scrollArea, 0, 0, 1, 1)
        self.gridLayout_10.addWidget(self.groupBox, 0, 2, 3, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.middle_widget)
        self.pushButton_2.setMinimumSize(QtCore.QSize(90, 50))
        self.pushButton_2.setMaximumSize(QtCore.QSize(81, 51))
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_10.addWidget(self.pushButton_2, 1, 1, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(20, 138, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_10.addItem(spacerItem1, 2, 1, 1, 1)
        self.tab_Widget = QtWidgets.QTabWidget(self.splitter)
        # self.tab_Widget.setStyleSheet("background-color:rgb(192,192,192);")
        self.tab_Widget.setObjectName("tab_Widget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.tab)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.refresh_pushButton = QtWidgets.QPushButton(self.tab)
        self.refresh_pushButton.setMinimumSize(QtCore.QSize(111, 24))
        self.refresh_pushButton.setMaximumSize(QtCore.QSize(111, 24))
        self.refresh_pushButton.setObjectName("refresh_pushButton")
        self.gridLayout_4.addWidget(self.refresh_pushButton, 1, 1, 1, 1)
        spacerItem2 = QtWidgets.QSpacerItem(129, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_4.addItem(spacerItem2, 1, 2, 1, 1)
        spacerItem3 = QtWidgets.QSpacerItem(680, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_4.addItem(spacerItem3, 1, 0, 1, 1)
        self.groupBoxing = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont("黑体")
        font.setPointSize(20)
        self.groupBoxing.setFont(font)
        self.groupBoxing.setObjectName("groupBoxing")
        self.gridLayout_13 = QtWidgets.QGridLayout(self.groupBoxing)
        self.gridLayout_13.setObjectName("gridLayout_13")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBoxing)
        self.tableWidget.setMinimumSize(QtCore.QSize(451, 201))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.gridLayout_13.addWidget(self.tableWidget, 0, 0, 1, 1)
        self.widget = QtWidgets.QWidget(self.groupBoxing)
        self.widget.setMinimumSize(QtCore.QSize(297, 224))
        self.widget.setMaximumSize(QtCore.QSize(297, 224))
        self.widget.setObjectName("widget")
        self.gridLayout_9 = QtWidgets.QGridLayout(self.widget)
        self.gridLayout_9.setObjectName("gridLayout_9")
        spacerItem4 = QtWidgets.QSpacerItem(74, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_9.addItem(spacerItem4, 0, 0, 1, 2)
        self.label = QtWidgets.QLabel(self.widget)
        self.label.setMinimumSize(QtCore.QSize(101, 51))
        self.label.setMaximumSize(QtCore.QSize(101, 51))
        font = QtGui.QFont("黑体")
        font.setPointSize(20)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setObjectName("label")
        self.gridLayout_9.addWidget(self.label, 0, 2, 1, 1)
        spacerItem5 = QtWidgets.QSpacerItem(75, 17, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_9.addItem(spacerItem5, 0, 3, 1, 2)
        spacerItem6 = QtWidgets.QSpacerItem(64, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_9.addItem(spacerItem6, 1, 0, 1, 1)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.mon_label = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        self.mon_label.setFont(font)
        self.mon_label.setObjectName("mon_label")
        self.horizontalLayout_3.addWidget(self.mon_label)
        self.mon_label_Edit = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(20)
        self.mon_label_Edit.setFont(font)
        self.mon_label_Edit.setAlignment(Qt.AlignCenter)
        self.mon_label_Edit.setObjectName("mon_label_Edit")
        self.horizontalLayout_3.addWidget(self.mon_label_Edit)
        self.gridLayout_9.addLayout(self.horizontalLayout_3, 1, 1, 1, 3)
        spacerItem7 = QtWidgets.QSpacerItem(65, 17, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_9.addItem(spacerItem7, 1, 4, 1, 1)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.mon_label_hello = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        self.mon_label_hello.setFont(font)
        self.mon_label_hello.setObjectName("mon_label_hello")
        self.gridLayout.addWidget(self.mon_label_hello, 0, 0, 1, 1)
        self.mon_label_hello_Edit = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        self.mon_label_hello_Edit.setFont(font)
        self.mon_label_hello_Edit.setAlignment(Qt.AlignCenter)
        self.mon_label_hello_Edit.setObjectName("mon_label_hello_Edit")
        self.gridLayout.addWidget(self.mon_label_hello_Edit, 0, 1, 1, 1)
        self.mon_label_reply = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        self.mon_label_reply.setFont(font)
        self.mon_label_reply.setObjectName("mon_label_reply")
        self.gridLayout.addWidget(self.mon_label_reply, 1, 0, 1, 1)
        self.mon_label_reply_Edit = QtWidgets.QLabel(self.widget)
        font = QtGui.QFont("黑体")
        font.setPointSize(15)
        self.mon_label_reply_Edit.setFont(font)
        self.mon_label_reply_Edit.setAlignment(Qt.AlignCenter)
        self.mon_label_reply_Edit.setObjectName("mon_label_reply_Edit")
        self.gridLayout.addWidget(self.mon_label_reply_Edit, 1, 1, 1, 1)
        self.gridLayout_9.addLayout(self.gridLayout, 2, 0, 1, 5)
        spacerItem8 = QtWidgets.QSpacerItem(17, 44, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_9.addItem(spacerItem8, 3, 2, 1, 1)
        self.gridLayout_13.addWidget(self.widget, 0, 1, 1, 1)
        self.gridLayout_4.addWidget(self.groupBoxing, 0, 0, 1, 3)
        self.tab_Widget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.tab_2)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.groupBoxed = QtWidgets.QGroupBox(self.tab_2)
        font = QtGui.QFont("黑体")
        font.setPointSize(20)
        self.groupBoxed.setFont(font)
        self.groupBoxed.setObjectName("groupBoxed")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.groupBoxed)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.historyTableWidget = QtWidgets.QTableWidget(self.groupBoxed)
        self.historyTableWidget.setObjectName("historyTableWidget")
        self.historyTableWidget.setColumnCount(0)
        self.historyTableWidget.setRowCount(0)
        self.gridLayout_2.addWidget(self.historyTableWidget, 0, 0, 1, 1)
        self.gridLayout_5.addWidget(self.groupBoxed, 0, 0, 1, 1)
        self.tab_Widget.addTab(self.tab_2, "")
        self.gridLayout_8.addWidget(self.splitter, 1, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi()
        self.tab_Widget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(self)
        try:
            with open("dy.txt", "r") as f:
                text = f.read()
                text = text.split("#")[:-1]
                for t in text:
                    self._addcomboBox(t)
        except Exception as e:
            # print(e)
            pass


        # 创建Excel对象,注意这里 必须传拷贝 self._comboBoxText.copy()
        self._excel = MyExcel(self._comboBoxText.copy())
        self.excelShowNamelabel.setText(self._excel.getEName())
        self.cecelShowPos_label.setText(self._excel.getEPath())

        # 创建TYpe类对象
        self._ty = TYpe(self,self._excel)
        # 表头
        self._tableWidgetHead = ["日期"]
        self._tableWidgetHead.extend(self._excel.getTwoList())
        # 创建添加账号obj
        self._cdy = CdyNumber(self,self._excel)
        # self._ty.mainobj(self)
        # 初始化
        self._init()

        # 增加号按钮
        self.serialNum_addPush.clicked.connect(lambda :self.test())
        # 下拉框事件
        self.serialNum_comboBox.currentIndexChanged.connect(lambda :self.k())
        # 点击打招呼次数
        self.hello_pushButton.clicked.connect(lambda :self.hello())
        # 点击回复次数次数
        self.reply_pushButton.clicked.connect(lambda :self.reply())
        # 招呼减
        self.hello_pushButton_2.clicked.connect(lambda :self.hello_())
        # 点击减
        self.reply_pushButton_2.clicked.connect(lambda :self.reply_())
        # 日历
        self.calendarWidget.clicked[QDate].connect(self.calendar)
        # 重置日历
        self.init_pushButton.clicked.connect(lambda :self._reset())
        # 刷新
        self.refresh_pushButton.clicked.connect(lambda :self.refresh())
        # 下载
        self.pushButton.clicked.connect(lambda :self._down())
        # 创建
        self.pushButton_2.clicked.connect(lambda :self._crate())
        # 下拉列表点击事件
        self.tableWidget.cellClicked[int,int].connect(self._duble)
        self.tableWidget.cellChanged[int,int].connect(self._tt)

    # 判断字符串是否为存数字
    def is_number(self,num):
        try:
            int(num)
            return True
        except Exception:
            return False

    # 获取当前操作的抖音号
    def _getdy(self):
        return self.serialNum_comboBox.currentText()

    # 修改表格当前位置的值
    def _tt(self,x,y):
        title = self._getdy()
        newx = x+4
        newy =y + self._comboBoxText.index(title) * len(self._excel.getTwoList()) + 1

        # 修改后的值
        newv =  self.tableWidget.item(x, y).text()
        # 判断self._tabValueing是否为存数字
        if self.is_number(self._tabValueing):
            # 在判断newv是否为数字
            if self.is_number(newv):
                self._excel.set_xy_value(newx,newy,int(newv))
            else:
                QMessageBox.information(self, "提示", "只能为[数字]")
                # 恢复原来的值
                c = QTableWidgetItem(str(self._tabValueing))
                c.setFont(QFont("黑体", 12, QFont.Bold))
                c.setForeground(QBrush(QColor(65, 105, 225)))
                # 设置居中
                c.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                self.tableWidget.setItem(x,y,c)
        elif type(self._tabValueing) == str:
            if self._tabValueing in ["是","否"]:
                if newv in ["是", "否"]:
                    self._excel.set_xy_value(newx, newy, newv)
                else:
                    QMessageBox.information(self, "提示", "只能[是]或者[否]")
                    # 恢复原来的值
                    c = QTableWidgetItem(str(self._tabValueing))
                    c.setFont(QFont("黑体", 12, QFont.Bold))
                    c.setForeground(QBrush(QColor(65, 105, 225)))
                    # 设置居中
                    c.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget.setItem(x, y, c)
            else:
                self._excel.set_xy_value(newx, newy, newv)
        # print("oldv",self._tabValueing)
        # print("new",)

    # 获取表格当前位置的值
    def _duble(self,x,y):
        # 当前操作打抖音号
        title = self._getdy()
        x += 4
        y += self._comboBoxText.index(title)*len(self._excel.getTwoList())+1
        # tab当前值
        self._tabValueing = self._excel.get_xy_value(x,y)
        # print(self._excel.get_xy_value(x,y))

    # 初始化
    def _init(self):
        # 当前操作打抖音号
        title = self._getdy()
        # print(title)
        self._excel.setDY(title)  # 设置抖音号

        self.groupBoxing.setTitle("预览当前抖音号[{}]".format(title))
        self.groupBoxed.setTitle("操作历史")

        # QSS
        self._qss()
        # 历史
        self._history()
        # 当前预览
        self.ExcelPreviewING()
        # 统计数据显示
        self.dataShow()

    # QSS
    def _qss(self):
        # 主页背景
        self.setStyleSheet("background-color:rgb(245,245,245);")
        # 下拉框
        self.serialNum_comboBox.setStyleSheet(QSS.serialNum_comboBoxColor())
        # 增加号
        self.serialNum_addPush.setStyleSheet(QSS.serialNum_addPushColor())
        # 当前Excel
        self.excelName_label.setStyleSheet(QSS.excelName_labelColor())
        # 位置
        self.excelPos_label.setStyleSheet(QSS.excelPos_labelColor())
        # 名称
        self.excelShowNamelabel.setStyleSheet(QSS.cecelShowPos_labelColor())
        # 路径
        self.cecelShowPos_label.setStyleSheet(QSS.cecelShowPos_labelColor())
        # 刷新
        self.refresh_pushButton.setStyleSheet(QSS.refresh_pushButtonColor())
        # 下载
        self.pushButton.setStyleSheet(QSS.pushButtonColor())
        # 日历
        # self.calendarWidget.setStyleSheet("background-color:rgb(192,192,192);")
        # 重置按钮
        self.init_pushButton.setStyleSheet(QSS.init_pushButtonColr())
        # 创建按钮
        self.pushButton_2.setStyleSheet(QSS.pushButton_2Colr())
        # 控件左
        self.hello_pushButton.setStyleSheet(QSS.hello_pushButtonColor())
        self.reply_pushButton.setStyleSheet(QSS.hello_pushButtonColor())
        # 控件右
        self.hello_pushButton_2.setStyleSheet(QSS.hello_pushButton_2Color())
        self.reply_pushButton_2.setStyleSheet(QSS.hello_pushButton_2Color())
        # 当前月
        self.mon_label_Edit.setStyleSheet(QSS.mon_label_EditColor())

    # 创建
    def _crate(self):
        self._ty.show()

    # 下载
    def _down(self):
        file = self._excel.getEPath() + "/" + self._excel.getEName()
        # directory=QFileDialog.getOpenFileName(self, "选择文件", os.getcwd())
        directory=QFileDialog.getSaveFileName(self,"save","All Files (*)")
        path = directory[0]
        # print(directory)
        # fileName = directory[0].split("/")[-1]
        try:
            if path:
                path += ".xlsx"
                self._excel.downExcel(path)
                QMessageBox.information(self, "提示", "下载完成")
        except Exception as e:
            # print(e)
            pass
    # 刷新
    def refresh(self):
        title = self.serialNum_comboBox.currentText()
        # print(title)
        self._excel.setDY(title)  # 设置抖音号
        self.groupBoxing.setTitle("预览当前抖音号[{}]".format(title))
        self.groupBoxed.setTitle("预览当前抖音号[{}]".format(title))
        # 当前预览
        self.ExcelPreviewING()
        # 统计数据显示
        self.dataShow()

    # 重置事件  重置日历
    def _reset(self):
        _year, _mon, _day = time.strftime("%Y_%m_%d", time.localtime()).split("_")
        self.calendarWidget.setSelectedDate(QDate(int(_year), int(_mon), int(_day)))

    # 日历事件
    def calendar(self,date):
        # print(date.toString("yyyy-MM-dd"))
        y,m,d = date.toString("yyyy-MM-dd").split("-")
        if d[0] == "0": # 去0
            d = d[-1]
        item=self.tableWidget.findItems(d,Qt.MatchExactly)
        row = item[0]
        row.setBackground(QBrush(QColor(100,149,237)))  # 设置背景
        # row.setForeground(QBrush(QColor(255,0,0)))  # 设置前景

        # 获取每一小格的高度
        grid = 780//self._excel.mon()
        # 滑向日历选中项
        self.tableWidget.verticalScrollBar().setSliderPosition(grid*row.row())

    # 当前时间
    def _mytime(self):
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

    # 打招呼次数增
    def hello(self):
        self._excel.setTwoSignalDict("打招呼次数")
        self._excel.setBtnName("打招呼次数")
        self._excel.AExe()
        # 当前时间
        timeing = self._mytime()
        self._crateHistory("打招呼次数",timeing)
        # 当前预览
        # self.ExcelPreviewING()
        self._localShowRefresh()
        # 统计数据显示
        self.dataShow()

    # 回复次数增
    def reply(self):
        self._excel.setTwoSignalDict("回复次数")
        self._excel.setBtnName("回复次数")
        self._excel.AExe()
        # 当前时间
        timeing = self._mytime()
        self._crateHistory("回复次数", timeing)
        # 当前预览
        # self.ExcelPreviewING()
        self._localShowRefresh()
        # 统计数据显示
        self.dataShow()

    # 打招呼次数减
    def hello_(self):
        self._excel.setTwoSignalDict("打招呼次数")
        self._excel.setBtnName("打招呼次数")
        self._excel.AExe(AR=False)
        # 当前时间
        timeing = self._mytime()
        self._crateHistory("打招呼次数",timeing)
        # 当前预览
        # self.ExcelPreviewING()
        self._localShowRefresh()
        # 统计数据显示
        self.dataShow()

    # 回复次数减
    def reply_(self):
        self._excel.setTwoSignalDict("回复次数")
        self._excel.setBtnName("回复次数")
        self._excel.AExe(AR=False)
        # 当前时间
        timeing = self._mytime()
        self._crateHistory("回复次数", timeing)
        # 当前预览
        # self.ExcelPreviewING()
        self._localShowRefresh()
        # 统计数据显示
        self.dataShow()


    # 下拉框事件
    def k(self):
        title = self.serialNum_comboBox.currentText()
        self._excel.setDY(title)  # 切换当前抖音号
        self.groupBoxing.setTitle("预览当前抖音号[{}]".format(title))
        self.groupBoxed.setTitle("操作历史")
        # 当前预览
        self.ExcelPreviewING()
        # 统计数据显示
        self.dataShow()

    # 增加号按钮
    def test(self):
        self._cdy.show()


    # 增加抖音号(下拉框)
    def _addcomboBox(self,dy:str):
        self._comboBoxText.append(dy)
        # print("lllllll:",self._comboBoxText)
        # 同步到Excel
        try:
            if not self._excel.getDyList():
                self._excel.addDYList(dy)
            elif self._excel.getDyList()[0] == "demo":
                self._excel.setDYList([dy])
            else:
                self._excel.addDYList(dy)
        except Exception as e:
            # print("错误001:",e)
            pass
        # 下拉框显示同步
        self.serialNum_comboBox.addItem("")
        self.serialNum_comboBox.setItemText(self._comboBoxIndex,dy)
        self.serialNum_comboBox.setCurrentIndex(self._comboBoxIndex)  # 设置当前项为选择项
        # 加一
        self._comboBoxIndex += 1

    # 删除抖音号(下拉框)
    def _delcomBox(self,dy:str):
        self._comboBoxText.remove(dy)
        # 同步到Excel
        try:
            if self._excel.getDyList():
                self._excel.delDYList(dy)
        except Exception:
            pass
        # 下拉框显示同步
        self.serialNum_comboBox.clear()  # 清空
        for i in range(len(self._comboBoxText)):  # 在填充
            self.serialNum_comboBox.addItem("")
            self.serialNum_comboBox.setItemText(i, self._comboBoxText[i])
        # 减一
        self._comboBoxIndex -= 1

    # 历史tableWidget的创建
    def _history(self):
        self.historyTableWidget.setRowCount(1000)
        self.historyTableWidget.setColumnCount(2)
        self.historyTableWidget.setHorizontalHeaderLabels(["时间", "操作"])
        # 隐藏头
        self.historyTableWidget.verticalHeader().setVisible(False)

    # 创建历史
    def _crateHistory(self,btnName,timeStr:str):
        ch = QTableWidgetItem(timeStr)
        hello = QTableWidgetItem(btnName)
        # 设置居中
        hello.setTextAlignment(Qt.AlignHCenter|Qt.AlignVCenter)
        ch.setTextAlignment(Qt.AlignHCenter|Qt.AlignVCenter)

        self.historyTableWidget.setItem(self._historyIndex, 0, ch)
        self.historyTableWidget.setItem(self._historyIndex, 1, hello)

        self._historyIndex += 1

    # comBox创建月
    def _crateMon(self,mon):
        # 生成月
        for i in range(1, mon + 1):
            c = QTableWidgetItem(str(i))
            c.setFont(QFont("黑体", 12, QFont.Bold))
            # 设置居中
            c.setTextAlignment(Qt.AlignHCenter|Qt.AlignVCenter)
            self.tableWidget.setItem(i - 1, 0, c)
        # 设置表头标签
        self.tableWidget.setHorizontalHeaderLabels(self._tableWidgetHead)

    #   Excel预览
    def ExcelPreview(self):
        mon = self._excel.mon()
        self.tableWidget.setRowCount(mon)
        # 重新创建头
        self._tableWidgetHead = [self._tableWidgetHead[0]]
        self._tableWidgetHead.extend(self._excel.getTwoList())


        self.tableWidget.setColumnCount(len(self._tableWidgetHead))
        # 设置表头标签
        self.tableWidget.setHorizontalHeaderLabels(self._tableWidgetHead)
        self.tableWidget.horizontalHeader().setFont(QFont("黑体", 14, QFont.Bold))
        # self.tableWidget.horizontalHeader().setStyleSheet("background-color:red")
        self.tableWidget.horizontalHeader().setStyleSheet("color:rgb(0,0,0)")
        # 隐藏头
        self.tableWidget.verticalHeader().setVisible(False)
        # 生成月
        self._crateMon(mon)

    # 数据显示局部刷新
    def _localShowRefresh(self):
        name = self._excel.getBtnName()

        col = self._excel.getTwoList().index(name)+1
        row = int(time.strftime("%d", time.localtime()))-1

        c = QTableWidgetItem(str(self._excel.getData()))
        c.setFont(QFont("黑体", 12, QFont.Bold))
        c.setForeground(QBrush(QColor(65, 105, 225)))
        # 设置居中
        c.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.tableWidget.setItem(row, col, c)

    # 当前预览
    def ExcelPreviewING(self):
        self.ExcelPreview()
        # print("--------------------------------------------")
        self._excel.oldDataSyn()
        data = self._excel.parsingExcelDict()
        # print("data:",data)
        # 当前抖音号
        title = self.serialNum_comboBox.currentText()
        # 清空,生成月
        self.tableWidget.clear()
        self._crateMon(self._excel.mon())
        TwoList = self._excel.getTwoList() # 标签表
        # print("T:",TwoList)
        try:
            for k, v in data[title].items():
                intk = int(k) - 3
                for pk, pv in data[title][k].items():
                    c = QTableWidgetItem(str(pv))
                    c.setFont(QFont("黑体", 12, QFont.Bold))
                    c.setForeground(QBrush(QColor(65,105,225)))
                    # 设置居中
                    c.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    if pk in TwoList:
                        self.tableWidget.setItem(intk - 1, TwoList.index(pk)+1, c)
        except Exception as e:
            # print("错误:",e)
            self.tableWidget.clear()
            self._crateMon(self._excel.mon())

    # 设置统计的数据
    def dataShow(self):
        title = self.serialNum_comboBox.currentText()
        # 设置月的天数
        self.mon_label_Edit.setText(str(self._excel.mon()))
        # 数据
        self._data = self._excel.helloReplyData()
        if title in self._data:
            self.mon_label_hello_Edit.setText(str(self._data[title]["打招呼次数"]))
            self.mon_label_reply_Edit.setText(str(self._data[title]["回复次数"]))
        else:
            self.mon_label_hello_Edit.setText("0")

            self.mon_label_reply_Edit.setText("0")


    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "ADD"))
        self.serialNum_label.setText(_translate("MainWindow", "抖音号"))
        self.serialNum_addPush.setText(_translate("MainWindow", "增加号"))
        self.excelName_label.setText(_translate("MainWindow", "当前EXcel:"))
        self.excelShowNamelabel.setText(_translate("MainWindow", "xxxxxxxxxxxxxxxxxxxxxxx"))
        self.excelPos_label.setText(_translate("MainWindow", "位        置:"))
        self.cecelShowPos_label.setText(_translate("MainWindow", "xxxxxxxxxxxxxxxxxxxxxxx"))
        self.pushButton.setText(_translate("MainWindow", "⬇️"))
        self.init_pushButton.setText(_translate("MainWindow", "重置日历"))
        self.groupBox.setTitle(_translate("MainWindow", "控件区"))
        self.reply_pushButton_2.setText(_translate("MainWindow", "回复减"))
        self.reply_pushButton.setText(_translate("MainWindow", "回复增"))
        self.hello_pushButton.setText(_translate("MainWindow", "招呼加"))
        self.hello_pushButton_2.setText(_translate("MainWindow", "招呼减"))
        self.pushButton_2.setText(_translate("MainWindow", "创建"))
        self.refresh_pushButton.setText(_translate("MainWindow", "刷新"))
        self.groupBoxing.setTitle(_translate("MainWindow", "预览[当前1号]"))
        self.label.setText(_translate("MainWindow", "统计数据"))
        self.mon_label.setText(_translate("MainWindow", "当月天数:"))
        self.mon_label_Edit.setText(_translate("MainWindow", "31"))
        self.mon_label_hello.setText(_translate("MainWindow", "当月打招呼总次数:"))
        self.mon_label_hello_Edit.setText(_translate("MainWindow", "7000"))
        self.mon_label_reply.setText(_translate("MainWindow", "当 月 回 复总次数:"))
        self.mon_label_reply_Edit.setText(_translate("MainWindow", "4000"))
        self.tab_Widget.setTabText(self.tab_Widget.indexOf(self.tab), _translate("MainWindow", "当前"))
        self.groupBoxed.setTitle(_translate("MainWindow", "预览[历史-当前1号]"))
        self.tab_Widget.setTabText(self.tab_Widget.indexOf(self.tab_2), _translate("MainWindow", "历史点击"))

        # 不可编辑
        # self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.historyTableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        # 克复制
        self.cecelShowPos_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        # 切割出 年月日
        self._reset()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    ui = MainUI()
    ui.show()

    sys.exit(app.exec_())