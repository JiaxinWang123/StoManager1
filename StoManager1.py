#### Imports ####
from tkinter import *
from PIL import ImageTk, Image
import cv2
import numpy as np
import glob
import random
import pandas as pd
import math
import os
from tkinter import filedialog
import statistics
import string
from openpyxl.workbook import Workbook
import matplotlib.pyplot as plt
import re
from pandas import Series
import time
from time import time, gmtime, strftime
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import csv
from PyQt5.QtGui import QPixmap
import os, shutil
import webbrowser

######

#### Main Window ####
class Ui_StoManager1(object):
    def __init__(self):
        super(Ui_StoManager1, self).__init__()
        self.setupUi(StoManager1)
        self.selected_image_index = 0
        self.img_num = 1
        self.folder =str()


    def setupUi(self, StoManager1):
        StoManager1.setObjectName("StoManager1")
        StoManager1.resize(1225, 847)
        StoManager1.setWindowIcon(QtGui.QIcon("StoManager.ico"))
        self.model = QtGui.QStandardItemModel(StoManager1)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(3)
        sizePolicy.setVerticalStretch(3)
        sizePolicy.setHeightForWidth(StoManager1.sizePolicy().hasHeightForWidth())
        StoManager1.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        StoManager1.setFont(font)
        StoManager1.setStyleSheet("font: 8pt \"Arial\";\n"
"background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:1 rgba(30, 141, 162, 235));")
        self.centralwidget = QtWidgets.QWidget(StoManager1)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setSpacing(7)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setContentsMargins(2, 2, 2, 2)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout()
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit.setFont(font)
        self.lineEdit.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit.setClearButtonEnabled(True)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.pushButton_8 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"")
        self.pushButton_8.setObjectName("pushButton_8")
        self.pushButton_8.clicked.connect(self.loadInputFolder)
        self.horizontalLayout.addWidget(self.pushButton_8)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_2.setClearButtonEnabled(True)
        self.lineEdit_2.setObjectName("lineEdit_2")
        
        
        self.input_image_path_ = self.lineEdit.text()
        self.output_image_path_ = self.lineEdit_2.text()
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_9.setObjectName("pushButton_9")
        self.pushButton_9.clicked.connect(self.loadOutputFolder)
        self.horizontalLayout_2.addWidget(self.pushButton_9)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_3.setClearButtonEnabled(True)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_4.addWidget(self.lineEdit_3)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_4.setClearButtonEnabled(True)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.horizontalLayout_5.addWidget(self.lineEdit_4)
        self.verticalLayout.addLayout(self.horizontalLayout_5)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_5.setClearButtonEnabled(True)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.horizontalLayout_8.addWidget(self.lineEdit_5)
        self.verticalLayout.addLayout(self.horizontalLayout_8)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.run_analyze)
        self.horizontalLayout_3.addWidget(self.pushButton_3)
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"")
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_6.clicked.connect(lambda: [self.show_info_messagebox_normalize(),self.NormalizeFileNames()])
        self.horizontalLayout_3.addWidget(self.pushButton_6)
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_7.clicked.connect(self.Stomata_with_no_groups)
        self.horizontalLayout_3.addWidget(self.pushButton_7)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.verticalLayout_8.addLayout(self.verticalLayout)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setEnabled(True)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        self.progressBar.setPalette(palette)
        self.progressBar.setStatusTip("")
        self.progressBar.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"color: rgb(255, 255, 255);")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(True)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_8.addWidget(self.progressBar)
        self.verticalLayout_4.addLayout(self.verticalLayout_8)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.pushButton_15 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_15.setFont(font)
        self.pushButton_15.setAcceptDrops(False)
        self.pushButton_15.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_15.setObjectName("pushButton_15")
        self.pushButton_15.clicked.connect(lambda: [self.clear_model(), self.loadCsv()])
        self.horizontalLayout_9.addWidget(self.pushButton_15)
        self.pushButton_16 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_16.setFont(font)
        self.pushButton_16.setAcceptDrops(False)
        self.pushButton_16.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_16.setObjectName("pushButton_16")
        self.pushButton_16.clicked.connect(lambda: [self.clear_model(), self.loadExl()])
        self.horizontalLayout_9.addWidget(self.pushButton_16)
        self.verticalLayout_2.addLayout(self.horizontalLayout_9)
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        self.tableView.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"\n"
"")
        self.tableView.setObjectName("tableView")
        self.verticalLayout_2.addWidget(self.tableView)
        self.verticalLayout_4.addLayout(self.verticalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout_4, 0, 0, 1, 1)
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.input_previous_image_clicked)
        self.horizontalLayout_6.addWidget(self.pushButton)
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(lambda: [self.check_input_Path(), self.show_info_messagebox_original_image()])
        self.horizontalLayout_6.addWidget(self.pushButton_4)
        self.pushButton_10 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_10.setFont(font)
        self.pushButton_10.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_10.setObjectName("pushButton_10")
        self.pushButton_10.clicked.connect(self.input_next_image_clicked)
        self.horizontalLayout_6.addWidget(self.pushButton_10)
        self.verticalLayout_3.addLayout(self.horizontalLayout_6)
        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.graphicsView.sizePolicy().hasHeightForWidth())
        self.graphicsView.setSizePolicy(sizePolicy)
        self.graphicsView.setStyleSheet("\n"
"background-color: rgb(255, 255, 255);")
        self.graphicsView.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.graphicsView.setResizeAnchor(QtWidgets.QGraphicsView.AnchorViewCenter)
        self.graphicsView.setObjectName("graphicsView")
        self.verticalLayout_3.addWidget(self.graphicsView)
        self.gridLayout_2.addLayout(self.verticalLayout_3, 0, 0, 1, 1)
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.output_previous_image_clicked)
        self.horizontalLayout_7.addWidget(self.pushButton_2)
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"")
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.clicked.connect(lambda:[self.check_output_Path(), self.show_info_messagebox_original_image()])
        self.horizontalLayout_7.addWidget(self.pushButton_5)
        self.pushButton_11 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_11.setFont(font)
        self.pushButton_11.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_11.setObjectName("pushButton_11")
        self.pushButton_11.clicked.connect(self.output_next_image_clicked)
        self.horizontalLayout_7.addWidget(self.pushButton_11)
        self.verticalLayout_7.addLayout(self.horizontalLayout_7)
        self.graphicsView_2 = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView_2.setEnabled(True)
        self.graphicsView_2.setStyleSheet("\n"
"\n"
"background-color: rgb(255, 255, 255);")
        self.graphicsView_2.setObjectName("graphicsView_2")
        self.verticalLayout_7.addWidget(self.graphicsView_2)
        self.gridLayout_2.addLayout(self.verticalLayout_7, 1, 0, 1, 1)
        self.gridLayout.addLayout(self.gridLayout_2, 0, 1, 1, 1)
        self.verticalLayout_5.addLayout(self.gridLayout)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(214, 214, 214);")
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_10.addWidget(self.label_3)
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.label_9.setFont(font)
        self.label_9.setStyleSheet("color:rgba(255,0,0,160);")
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_10.addWidget(self.label_9)
        self.verticalLayout_5.addLayout(self.horizontalLayout_10)
        self.gridLayout_3.addLayout(self.verticalLayout_5, 0, 0, 1, 1)
        StoManager1.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(StoManager1)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1321, 26))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuView = QtWidgets.QMenu(self.menubar)
        self.menuView.setObjectName("menuView")
        self.menuMeasure = QtWidgets.QMenu(self.menubar)
        self.menuMeasure.setObjectName("menuMeasure")
        self.menuAnalysis = QtWidgets.QMenu(self.menubar)
        self.menuAnalysis.setObjectName("menuAnalysis")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        StoManager1.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(StoManager1)
        self.statusbar.setObjectName("statusbar")
        StoManager1.setStatusBar(self.statusbar)
        self.actionOpen = QtWidgets.QAction(StoManager1)
        self.actionOpen.setObjectName("actionOpen")
        self.actionOpen.triggered.connect(self.loadInputFolder)
        self.actionCreate = QtWidgets.QAction(StoManager1)
        self.actionCreate.setObjectName("actionCreate")
        self.actionDelect = QtWidgets.QAction(StoManager1)
        self.actionDelect.setObjectName("actionDelect")
        self.actionWindows = QtWidgets.QAction(StoManager1)
        self.actionWindows.setObjectName("actionWindows")
        self.actionWindows.triggered.connect(self.check_input_Path)
        self.actionGroup_Analysis = QtWidgets.QAction(StoManager1)
        self.actionGroup_Analysis.setObjectName("actionGroup_Analysis")
        self.actionGroup_Analysis.triggered.connect(self.Stomata)
        self.actionNon_Group_Analysis = QtWidgets.QAction(StoManager1)
        self.actionNon_Group_Analysis.setObjectName("actionNon_Group_Analysis")
        self.actionNon_Group_Analysis.triggered.connect(self.Stomata_with_no_groups)
        self.actionOutput_Img = QtWidgets.QAction(StoManager1)
        self.actionOutput_Img.setObjectName("actionOutput_Img")
        self.actionLabel_Info = QtWidgets.QAction(StoManager1)
        self.actionLabel_Info.setObjectName("actionLabel_Info")
        self.actionAnalysis_Result = QtWidgets.QAction(StoManager1)
        self.actionAnalysis_Result.setObjectName("actionAnalysis_Result")
        self.actionAnalysis_Result.triggered.connect(self.loadCsv)
        self.actionStatistical_Results = QtWidgets.QAction(StoManager1)
        self.actionStatistical_Results.setObjectName("actionStatistical_Results")
        self.actionStatistical_Results.triggered.connect(self.loadExl)
        self.actionStomatal_area = QtWidgets.QAction(StoManager1)
        self.actionStomatal_area.setObjectName("actionStomatal_area")
        self.actionStomatal_orientation = QtWidgets.QAction(StoManager1)
        self.actionStomatal_orientation.setObjectName("actionStomatal_orientation")
        self.action = QtWidgets.QAction(StoManager1)
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.action.setFont(font)
        self.action.setObjectName("action")
        self.action.triggered.connect(self.web_link_guthub)
        self.actionH = QtWidgets.QAction(StoManager1)
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.actionH.setFont(font)
        self.actionH.setObjectName("actionH")
        self.actionH.triggered.connect(self.web_link_google_scholar)
        self.menuFile.addAction(self.actionOpen)
        self.menuFile.addAction(self.actionCreate)
        self.menuFile.addAction(self.actionDelect)
        self.menuView.addAction(self.actionWindows)
        self.menuView.addAction(self.actionOutput_Img)
        self.menuView.addAction(self.actionLabel_Info)
        self.menuView.addAction(self.actionAnalysis_Result)
        self.menuView.addAction(self.actionStatistical_Results)
        self.menuMeasure.addAction(self.actionStomatal_area)
        self.menuMeasure.addAction(self.actionStomatal_orientation)
        self.menuAnalysis.addAction(self.actionGroup_Analysis)
        self.menuAnalysis.addAction(self.actionNon_Group_Analysis)
        self.menuHelp.addAction(self.action)
        self.menuHelp.addAction(self.actionH)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuView.menuAction())
        self.menubar.addAction(self.menuMeasure.menuAction())
        self.menubar.addAction(self.menuAnalysis.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(StoManager1)
        QtCore.QMetaObject.connectSlotsByName(StoManager1)

    def retranslateUi(self, StoManager1):
        _translate = QtCore.QCoreApplication.translate
        StoManager1.setWindowTitle(_translate("StoManager1", "StoManager1.v.0308.23_For_Hardwoods"))
        self.label_3.setText(_translate("StoManager1", "© Jiaxin Wang,  For questions or requests @ coolwjx@foxmail.com; jw3994@msstate.edu"))
        self.lineEdit.setText(_translate("StoManager1", "Input path"))
        self.pushButton_8.setText(_translate("StoManager1", "Input"))
        self.lineEdit_2.setText(_translate("StoManager1", "Output path"))
        self.pushButton_9.setText(_translate("StoManager1", "Output"))
        self.lineEdit_3.setText(_translate("StoManager1", "Input image size, the default is: 2048,1536"))
        self.lineEdit_4.setText(_translate("StoManager1", "Input pixels in 0.1 mm, the default is: 476"))
        self.lineEdit_5.setText(_translate("StoManager1", "Confidence threshold for detection the default is 0.3"))
        self.pushButton_3.setText(_translate("StoManager1", "Start Process"))
        self.pushButton_6.setText(_translate("StoManager1", "Normalize File"))
        self.pushButton_7.setText(_translate("StoManager1", "Statistical Analysis"))
        self.pushButton_15.setText(_translate("StoManager1", "Preview exported table for single image"))
        self.pushButton_16.setText(_translate("StoManager1", "Preview exported statistics for all images"))
        self.label_9.setText(_translate("StoManager1", "💕   LU"))
        self.pushButton.setText(_translate("StoManager1", "<< View previous"))
        self.pushButton_4.setText(_translate("StoManager1", "View original images"))
        self.pushButton_10.setText(_translate("StoManager1", "View next >>"))
        self.pushButton_2.setText(_translate("StoManager1", "<< View previous"))
        self.pushButton_5.setText(_translate("StoManager1", "View detected images"))
        self.pushButton_11.setText(_translate("StoManager1", "View next >>"))
        self.menuFile.setTitle(_translate("StoManager1", "File"))
        self.menuView.setTitle(_translate("StoManager1", "View"))
        self.menuMeasure.setTitle(_translate("StoManager1", "Measure"))
        self.menuAnalysis.setTitle(_translate("StoManager1", "Analysis"))
        self.menuHelp.setTitle(_translate("StoManager1", "Help"))
        self.actionOpen.setText(_translate("StoManager1", "Open"))
        self.actionCreate.setText(_translate("StoManager1", "Create"))
        self.actionDelect.setText(_translate("StoManager1", "Delete"))
        self.actionWindows.setText(_translate("StoManager1", "Original Img"))
        self.actionGroup_Analysis.setText(_translate("StoManager1", "Group Analysis"))
        self.actionNon_Group_Analysis.setText(_translate("StoManager1", "Non-Group Analysis"))
        self.actionOutput_Img.setText(_translate("StoManager1", "Output Img"))
        self.actionLabel_Info.setText(_translate("StoManager1", "Label Info"))
        self.actionAnalysis_Result.setText(_translate("StoManager1", "Analysis Results"))
        self.actionStatistical_Results.setText(_translate("StoManager1", "Statistical Results"))
        self.actionStomatal_area.setText(_translate("StoManager1", "Stomatal area"))
        self.actionStomatal_orientation.setText(_translate("StoManager1", "Stomatal orientation"))
        self.action.setText(_translate("StoManager1", "Visit my Git Hub"))
        self.actionH.setText(_translate("StoManager1", "Visit my Google Scholar"))

    def check_input_Path(self):
        
        self.input_image_path_ = self.lineEdit.text()
                # get first image file and show it
        image_files = glob.glob(self.input_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.input_image_path_ + "/" + '*jpg')) is True:
            if len(image_files) > 0:
                self.selected_image_index = 0
            else:
                self.selected_image_index = -1

            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView.setScene(scene)
        else:
            self.messagebox_define_input_path()
            pass
    
    def check_output_Path(self):
        
        self.output_image_path_ = self.lineEdit_2.text()
                # get first image file and show it
        image_files_2 = glob.glob(self.output_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.output_image_path_ + "/" + '*jpg')) is True:
            if len(image_files_2) > 0:
                self.selected_image_index = 0
            else:
                self.selected_image_index = -1

            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files_2[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView_2.setScene(scene)   
        else:
            self.messagebox_define_output_path()
            pass  

# next image

    def input_next_image_clicked(self):
        # get first image file and show it
        
        self.input_image_path_ = self.lineEdit.text()
        image_files = glob.glob(self.input_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.input_image_path_ + "/" + '*jpg')) is True:

            if self.selected_image_index in range(len(image_files)):
                # If reached the last image, set index to start over
                if self.selected_image_index == len(image_files) - 1:
                    self.selected_image_index = 0
                else:
                    self.selected_image_index += 1
            else:
                self.selected_image_index -= 1
            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView.setScene(scene)
            #self.selected_image_index += 1
        else:
            self.messagebox_define_input_path()
            pass




#next labeled image
    def output_next_image_clicked(self):
        
        # get first image file and show it
        self.output_image_path_ = self.lineEdit_2.text()
        image_files = glob.glob(self.output_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.output_image_path_ + "/" + '*jpg')) is True:

            if self.selected_image_index in range(len(image_files)):
                # If reached the last image, set index to start over
                if self.selected_image_index == len(image_files) - 1:
                    self.selected_image_index = 0
                else:
                    self.selected_image_index += 1
            else:
                self.selected_image_index -= 1
            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView_2.setScene(scene)
            #self.selected_image_index += 1

        else:
            self.messagebox_define_output_path()
            pass


# previous image
    def input_previous_image_clicked(self):
        
        self.input_image_path_ = self.lineEdit.text()
        image_files = glob.glob(self.input_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.input_image_path_ + "/" + '*jpg')) is True:

            if self.selected_image_index == 0:
                self.selected_image_index = len(image_files) - 1
            else:
                self.selected_image_index -= 1
            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView.setScene(scene)
        else:
            self.messagebox_define_input_path()
            pass

# previous image
    def output_previous_image_clicked(self):
        
        self.output_image_path_ = self.lineEdit_2.text()
        image_files = glob.glob(self.output_image_path_ + "/" + '*jpg')
        if bool(glob.glob(self.output_image_path_ + "/" + '*jpg')) is True:

            if self.selected_image_index == 0:
                self.selected_image_index = len(image_files) - 1
            else:
                self.selected_image_index -= 1
            scene = QtWidgets.QGraphicsScene(StoManager1)
            pixmap = QPixmap(image_files[self.selected_image_index])
            item = QtWidgets.QGraphicsPixmapItem(pixmap)
            scene.addItem(item)
            self.graphicsView_2.setScene(scene)
        else:
            self.messagebox_define_output_path()
            pass


###### Main components ######

    # define a function to normalize the file names
    def NormalizeFileNames(self):
        self.input_image_path_ = self.lineEdit.text()
        Output_path = self.lineEdit_2.text()
        if bool(glob.glob(self.input_image_path_ + "/" + '*jpg')) is True:
            file_path = glob.glob(self.input_image_path_ + "/" + '*jpg')[2]
            file_name = file_path[len(self.input_image_path_) + 1:-4]  # Extract the file name from the file path
            file_name_split = file_name.split(",")
            Split_name = []
            for s in file_name_split:
                s = re.sub("\s+", ",", s.strip())
                Split_name.append(s)
            Split_name = ",".join(Split_name)
            Split_name= Split_name.split(",")
            print(Split_name)  # Split the site, block, and clone info from the file name

            if len(Split_name) >= 5:
                        
                Site = str()
                Block = str()
                Year = str()
                Month = str()
                Tail = str()
                Clone = str()
                file_names = str()

                # create list of Site, Block, Year, Month, Tail, and Clone here, or import those information from .txt files and read them as lists

                site = ["M","Mon","m","mon","P","p","Pon","pon","PON","M ","Mon ","m ","mon ","P ","p ","Pon ","pon ","PON "] # create a list containing all potential sites
                block = ["b1","b2","b5","B1","B2","B5","b1 ","b2 ","b5 ","B1 ","B2 ","B5 "," b1"," b2"," b5"," B1"," B2"," B5"] # create a list containing all potential blocks
                year = ["20","21","22","23","24"," 20"," 21"," 22"," 23"," 24","20 ","21 ","22 ","23 ","24 ","2022","2020","2021","2019"] # create a list containing all potential years
                month = ["June","june","July","july","Aug","aug","AUG","JULY","JUNE"," June"," june"," July"," july"," Aug"," aug",
                " AUG"," JULY"," JUNE","June ","june ","July ","july ","Aug ","aug ","AUG ","JULY ","JUNE ","Sept","Sep"," Sept","sep"," sep"] # create a list containing all potential months
                tail = ["1","2","3","4","5","6","7","8","9","10","Aci","Aci1","Aci2","Aci3","Aci4","Aci5","Aci6","Aci7","Aci8","Aci9",
                "Aci10","aci","aci1","aci2","aci3","aci4","aci5","aci6","aci7","aci8","aci9","aci10","ACi","aCI"," 1"," 2"," 3"," 4"," 5",
                " 6"," 7"," 8"," 9"," 10"," Aci"," Aci1"," Aci2"," Aci3"," Aci4"," Aci5"," Aci6"," Aci7"," Aci8"," Aci9"," Aci10"," aci",
                " aci1"," aci2"," aci3"," aci4"," aci5"," aci6"," aci7"," aci8"," aci9"," aci10"," ACi"," aCI","1 ","2 ","3 ","4 ","5 ","6 ",
                "7 ","8 ","9 ","10 ","Aci ","Aci1 ","Aci2 ","Aci3 ","Aci4 ","Aci5 ","Aci6 ","Aci7 ","Aci8 ","Aci9 ","Aci10 ","aci ","aci1 ",
                "aci2 ","aci3 ","aci4 ","aci5 ","aci6 ","aci7 ","aci8 ","aci9 ","aci10 ","ACi ","aCI ","40X","20X"," 40X"," 20x"," 40X ",
                " 20X ","40x ","20X ","EC","DN","DM","DD"] # create a list containing all potential tails
                clone = ["S7C2","s7c2","s7c4","S7C4","ST-66","st-66","st66","ST66","19","22","110412","11412","120-4","3-1","6-1","6-5","ST-70",
                "ST-75","st-70","st-75","st70","st75","6323","6329","8019","9225","9707","11690","13693","13724","24033","24056","24066","24114",
                "24120","24159","29310","433","11785","11789","11795","11797","11802","11840","13849","14278","14340","24250","7903","8729","10016",
                "24245","24301","24304","24326","24340","29262","29270","11666","9709","14492","11691","13700","14486","8015","8002","14507","13725",
                "11711","8852","10243","9225","14591","6323","13695","9711","14490","7388","9189","11732","11822","433","11867","11859","7389","8198",
                "8729","7903","10029","9552","9755","9980","8717","11733","111733","120-4","S7C20","s7c20","ST75","1BB-3","113B-3","113b-3","112107",
                "S13C20","s13c20","106B-1","106b-1","47-5","6-4","9198","7938"] # create a list containing all potential clones, you can also import clone.txt file and read it as a list

                for file_path in glob.glob(self.input_image_path_ + "/" + '*jpg'):
                    file_name = file_path[len(self.input_image_path_) + 1:-4]  # Extract the file name from the file path
                    file_name_split = file_name.split(",")
                    Split_name = []
                    for s in file_name_split:
                        s = re.sub("\s+", ",", s.strip())
                        Split_name.append(s)
                    Split_name = ",".join(Split_name)
                    Split_name= Split_name.split(",")
                    print(Split_name)  # Split the site, block, and clone info from the file name

                    if len(Split_name) >= 5:
                        for name in Split_name:

                            if name in site:
                                Site = name

                            elif name in block:
                                Block = name

                            elif name in year:
                                Year = name

                            elif name in month:
                                Month = name

                            elif name in tail:
                                Tail = name

                            elif name in clone:
                                Clone = name
                    elif len(Split_name) <=4:

                        file_names =",".join([str(item) for item in Split_name])


                    else:
                        for name in Split_name:
                            if name in site:
                                Site = name
                            elif name in block:
                                Block = name
                            elif name in year:
                                Year = name
                            elif name in month:
                                Month = name
                            elif name in clone:
                                Clone = name
                            
                    if len(Split_name) > 4:
                        New_file_name = Site + "," + Block + "," + Clone + "," + Month + "," + Year + "," + Tail + ".jpg"
                        if not os.path.exists(self.input_image_path_ + "/" + New_file_name):
                            os.rename(file_path, self.input_image_path_ + "/" + New_file_name)
                    else:
                        New_file_name = file_names + ".jpg"
                        if not os.path.exists(self.input_image_path_ + "/" + New_file_name):
                            os.rename(file_path, self.input_image_path_ + "/" + New_file_name)
            else:
                pass
        else:
            self.show_info_messagebox_normalize_2()
            pass        


    def run_analyze(self):  ## the main funtion with the yolov3 model we trained to detect our stomata
        
        confidence_set = self.lineEdit_5.text()
        self.input_image_path_ = self.lineEdit.text()
        Output_path = self.lineEdit_2.text()

        Input_img_size = self.lineEdit_3.text()
        Input_pixels = self.lineEdit_4.text()
        if bool(glob.glob(self.input_image_path_ + "/" + '*jpg')) is True:

            self.img_num = 1

            if Input_img_size =="Input image size, the default is: 2048,1536" or Input_img_size =="" or Input_img_size == " ":
                Input_img_size = "2048,1536"
            else: 
                Input_img_size = self.lineEdit_3.text()

            if Input_pixels =="Input pixels in 0.1 mm, the default is: 476" or Input_pixels =="" or Input_pixels ==" ":
                Input_pixels = "476"
            else: 
                Input_pixels = self.lineEdit_4.text()
                        

            split_img_size = Input_img_size.split(",")

            pixle = float(Input_pixels)

            Img_area = float(split_img_size[0])*float(split_img_size[1])
            
            whole_img_area = Img_area/((pixle)*10.0)**2

            # Load Yolo
            net = cv2.dnn.readNet("yolov3_training_last.weights", "yolov3_testing.cfg")

            # Name custom object
            classes = ["whole_stomata", "stomata"]

            # Images path
            # image_path_input = input("Please type your image path:")
            image_path_input = self.input_image_path_
            images_path = glob.glob(str(image_path_input) + r"\*.jpg")

            print(images_path)
            layer_names = net.getLayerNames()
            output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]
            colors = np.random.uniform(0, 100, size=(len(classes), 3))
            # Insert here the path of your images
            random.shuffle(images_path)
            # loop through all the images

            save_files = []
            if bool(save_files) == False:
                # save_file = input('Please type the path where you want to save the export csv and jpg files:')
                save_file = Output_path
                save_files.append(save_file)
                save_file_str = ''.join(map(str, save_files))
            

            for img_path in images_path:
                # Loading image
                confidence_set = self.lineEdit_5.text()
            
                if confidence_set =="Confidence threshold for detection the default is 0.3" or confidence_set =="" or confidence_set ==" ":
                    confidence_set = 0.3
                else: 
                    confidence_set = float(self.lineEdit_5.text())
                
                self.progressBar.setValue(int((self.img_num / len(images_path)*100)))
                QtWidgets.QApplication.processEvents()

                self.img_num +=1
                
                img = cv2.imread(img_path)
                # img = cv2.resize(img, None, fx=0.5, fy=0.5)
                height, width, channels = img.shape

                # Detecting objects
                blob = cv2.dnn.blobFromImage(img, 0.00392, (416, 416), (0, 0, 0), True, crop=False)

                net.setInput(blob)
                outs = net.forward(output_layers)
                # Showing information on the screen
                class_ids = []
                class_ids_2 = []
                centers = []
                confidences = []
                boxes = []
                number_of_whole_stomata = []
                number_of_stomata = []
                list_of_width = []
                list_of_height = []
                image_paths = []
                orientations = []
                labels = []
                list_of_image_width = []
                list_of_image_height = []
                list_of_all_stomata_areas = []
                list_of_whole_stomatal_area = []
                list_of_x = []
                list_of_y = []
                Whole_stomata_density = []
                for out in outs:
                    for detection in out:
                        scores = detection[5:]
                        class_id = np.argmax(scores)
                        confidence = scores[class_id]
                        if confidence > confidence_set:
                            # Object detected

                            center_x = (detection[0] * width)
                            center_y = (detection[1] * height)
                            w = (detection[2] * width)
                            h = (detection[3] * height)
                            # Rectangle coordinates
                            x = (center_x - w / 2)
                            y = (center_y - h / 2)

                            boxes.append([x, y, w, h])
                            centers.append([center_x, center_y])
                            confidences.append(float(confidence))
                            class_ids.append(class_id)
                # put class id, confidence, x, y, width, height to a txt file
                
                indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.1, 0.4)  

                # print(len(indexes[class_id]))
                font = cv2.FONT_HERSHEY_PLAIN
                for i in range(len(boxes)):

                    if i in indexes:
                        x, y, w, h = boxes[i]
                        label = str(classes[class_ids[i]])
                        color = colors[class_ids[i]]
                        
                        cv2.rectangle(img, (int(x), int(y)), (int(x) + int(w), int(y) + int(h)), color, 1)
                        cv2.putText(img, label, (int(x), int(y) + 42), font, 1, color, 1)
                        cv2.putText(img, str(int(h)), (int(x), int(y) + 60), font, 1, color, 1)
                        cv2.putText(img, str(int(w)), (int(x), int(y) + 28), font, 1, color, 1)
                        cv2.putText(img, str(round(confidences[i], 2)), (int(x), int(y) + 12), font, 1, color, 1)

                        # cv2.putText(img, str(confidence), (x, y - 5), font, 1, color, 1)
                        if label == "whole_stomata":
                            number_of_whole_stomata.append(class_ids[i])
                            list_of_all_stomata_areas.append(
                                (w * h) * 0.6878 + 806)  ## Build linear regression model for adjusted area of whole_stomata
                            list_of_whole_stomatal_area.append((w * h) * 0.6878 + 806)
                        elif label == "stomata":  ## Build linear regression model for adjusted area of stomata
                            number_of_stomata.append(class_ids[i])
                            list_of_all_stomata_areas.append(
                                (w * h + 116.08) / 1.7684)
                            
                        orientation = math.log(w/h)*(-92.2325)+44.5222
                        orientations.append(orientation)

                        list_of_width.append(w)
                        list_of_height.append(h)
                        list_of_x.append(x)
                        list_of_y.append(y)
                        image_paths.append(img_path)
                        labels.append(label)
                        list_of_image_height.append(height)
                        list_of_image_width.append(width)
                        class_ids_2.append(class_ids[i])
                
                with open(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".txt", 'w') as f:
                    
                    for m in range(len(list_of_width)):
                        w_2 = list_of_width[m]/width
                        h_2 = list_of_height[m]/height
                        x_center = (list_of_x[m]+list_of_width[m]/2)/width
                        y_center = (list_of_y[m]+list_of_height[m]/2)/height


                        x = (center_x - w / 2)

                        f.write(str((str(class_ids_2[m]) + " " + str(confidences[m])+ " "+ str(x_center) + " " + str(y_center) + " " + str(w_2) + " " + str(h_2))))
                        f.write('\n')  

                list_of_whole_stomatal_area_ratio = sum(list_of_whole_stomatal_area) / (
                        statistics.mean(list_of_image_width) * statistics.mean(list_of_image_height))
                heights = list_of_height
                widths = list_of_width
                image_name = img_path
                image_width = list_of_image_width
                image_height = list_of_image_height
                all_stomata_areas = [(float(k)/Img_area*1000000) for k in list_of_all_stomata_areas]
                Whole_stomata_density=number_of_whole_stomata
                image_name = {"Labels": labels, "Width_(pixels)": widths, "Height_(pixels)": heights,
                            "Orientation": orientations,
                            "Num_of_Stomata": len(number_of_stomata), "Num_of_whole_stomata": len(number_of_whole_stomata),
                            "Width_of_image_(pixels)": image_width, "Height_of_image_(pixels)": image_height,
                            "All_sotmata_area_(mum2)": all_stomata_areas,
                            "Whole_stomata_area_ratio": list_of_whole_stomatal_area_ratio,
                            "Whole_stomata_density": len(Whole_stomata_density)/whole_img_area}

                df = pd.DataFrame(image_name)

                # print(image_name)
                # df.to_csv('width_height.csv')
                img_path = str(img_path)
                new_img_path = img_path.replace("\\", "")
                new_img_path = img_path.replace(".jpg", "")
                print(new_img_path)

                file = df.to_csv(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".csv", sep=',', index=True)
                cv2.imwrite(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".jpg", img)

            self.show_info_messagebox()
        else:
            self.messagebox_define_input_path_2()
            pass

    # Define a function to analyze the  stomatal data

    def Stomata(self):
        # Create empty lists to hold the values that we are going to extract
        self.output_image_path_ = self.lineEdit_2.text()
        self.file_path = glob.glob(self.output_image_path_ + '/' + '*.csv')[1]
        if bool(glob.glob(self.output_image_path_ + "/" + '*csv')) is True:
            


            WST_Area_median = []
            WST_Area_max = []
            WST_Area_min = []
            WST_Area_mean = []
            WST_Area_var = []
            WST_Area_std = []
            WST_Number = []
            WST_Area_Ratio = []
            WST_Orientation = []
            WST_Orientation_var = []
            WST_Density = []
            WST_Area = []
            Site = []
            Block = []
            Clone = []
            Month = []
            Year = []

            # Using for loop to go through all csv files
            if bool(glob.glob(self.output_image_path_ + '/' + 'Statistics.csv')) is True:
                os.remove(glob.glob(self.output_image_path_ + '/' + 'Statistics.csv')[0])
            else:
                self.img_num =1
                # for file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                self.file_path = glob.glob(self.output_image_path_ + '/' + '*.csv')[1]
                single_csv_file = pd.read_csv(self.file_path, low_memory=False)  # Read the csv and assign it to a variable
                file_name = self.file_path[len(self.output_image_path_) + 1:-4]  # Extract the file name from the file path
                Split_name = file_name.split(",")                 
                if len(Split_name)>=5:
                                    
                    for self.file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                        single_csv_file = pd.read_csv(self.file_path, low_memory=False)  # Read the csv and assign it to a variable
                        file_name = self.file_path[len(self.output_image_path_) + 1:-4]  # Extract the file name from the file path
                        Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
                        
                        Site.append(Split_name[0])  # Add the site to the Site list
                        Block.append(Split_name[1])
                        Clone.append(Split_name[2])
                        Month.append(Split_name[3])
                        Year.append(Split_name[4])

                        # Extract the parameters from each single csv files

                        Orientation = single_csv_file["Orientation"]  # Extract the Orientation
                        Whole_stomata_number = single_csv_file["Num_of_whole_stomata"]  # Extract the number of whole stomata
                        Whole_stomata_areas = single_csv_file[single_csv_file['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
                        Whole_stomata_area_ratio = single_csv_file['Whole_stomata_area_ratio']
                        Whole_stomata_density = single_csv_file["Whole_stomata_density"]

                        # Remove the outliers before calculating orientation variance
                        Q1 = np.percentile(Orientation, 25, method = 'midpoint')
                        Q3 = np.percentile(Orientation, 75,method = 'midpoint')

                        IQR = Q3 - Q1

                        upper = np.where(Orientation >= (Q3+1.5*IQR))
                        lower = np.where(Orientation <= (Q1-1.5*IQR))

                        Orientation.drop(upper[0], inplace = True)
                        Orientation.drop(lower[0], inplace = True)


                        # calculate the median, mean, variance of each leaf stomata number, stomata area

                        Whole_stomata_areas_median = np.median(Whole_stomata_areas)
                        Whole_stomata_areas_max = np.max(Whole_stomata_areas)
                        Whole_stomata_areas_min = np.min(Whole_stomata_areas)
                        Whole_stomata_areas_mean = np.mean(Whole_stomata_areas)
                        Whole_stomata_areas_var = np.var(Whole_stomata_areas)
                        Whole_stomata_areas_std = np.std(Whole_stomata_areas)

                        WST_Area_median.append(Whole_stomata_areas_median)
                        WST_Area_max.append(Whole_stomata_areas_max)
                        WST_Area_min.append(Whole_stomata_areas_min)
                        WST_Area_mean.append(Whole_stomata_areas_mean)
                        WST_Area_var.append(Whole_stomata_areas_var)
                        WST_Area_std.append(Whole_stomata_areas_std)
                        WST_Orientation.append(np.median(Orientation))
                        WST_Orientation_var.append(np.var(Orientation))

                        # Extract the number of stomata
                        WST_Density.append(np.mean(Whole_stomata_density))
                        WST_Number.append(np.mean(Whole_stomata_number))
                        WST_Area.append(np.mean(Whole_stomata_areas))
                        WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))
                        self.progressBar.setValue(int((self.img_num / len(glob.glob(self.output_image_path_ + '/' + '*.csv'))*100)))
                        QtWidgets.QApplication.processEvents()

                        self.img_num +=1
                # else:

                #     pass

                # Convert all lists into pd.series

                Site = pd.Series(Site, dtype=pd.StringDtype(), name="Site")
                Block = pd.Series(Block, dtype=pd.StringDtype(), name="Block")
                Clone = pd.Series(Clone, dtype=pd.StringDtype(), name="Clone")
                Month = pd.Series(Month, dtype=pd.StringDtype(), name="Month")
                Year = pd.Series(Year, dtype=pd.StringDtype(), name="Year")
                WST_Number = pd.Series(WST_Number, dtype=pd.Float64Dtype(), name="WST_Number")
                WST_Area_Ratio = pd.Series(WST_Area_Ratio, dtype=pd.Float64Dtype(), name="WST_Area_Ratio")
                WST_Area_median = pd.Series(WST_Area_median, dtype=pd.Float64Dtype(), name="WST_Area_median")
                WST_Area_max = pd.Series(WST_Area_max, dtype=pd.Float64Dtype(), name="WST_Area_max")
                WST_Area_min = pd.Series(WST_Area_min, dtype=pd.Float64Dtype(), name="WST_Area_min")
                WST_Area_mean = pd.Series(WST_Area_mean, dtype=pd.Float64Dtype(), name="WST_Area_mean")
                WST_Area_var = pd.Series(WST_Area_var, dtype=pd.Float64Dtype(), name="WST_Area_var")
                WST_Area_std = pd.Series(WST_Area_std, dtype=pd.Float64Dtype(), name="WST_Area_std")
                WST_Density = pd.Series(WST_Density, dtype=pd.Float64Dtype(), name ="WST_Density")
                WST_Orientation = pd.Series(WST_Orientation, dtype=pd.Float64Dtype(), name ="WST_Orientation")
                WST_Orientation_var = pd.Series(WST_Orientation_var, dtype=pd.Float64Dtype(), name ="WST_Orientation_var")

                # Put all extracted parameters into a data frame
                Output_data = pd.concat(
                    [Site, Block, Clone, Month, Year, WST_Number, WST_Area_Ratio,WST_Density, WST_Area_median, WST_Area_max, WST_Area_min,
                    WST_Area_mean, WST_Area_var, WST_Area_std, WST_Orientation, WST_Orientation_var], axis=1)
                Output_data = pd.DataFrame(data=Output_data)
                random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
                Output_data.to_excel(self.output_image_path_ + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
                self.show_info_messagebox_group_analysis1()

        else:
            self.show_info_messagebox_group_analysis_2()
            pass    

    def Stomata_with_no_groups(self):
        # Create empty lists to hold the values that we are going to extract
        self.output_image_path_ = self.lineEdit_2.text()
        if bool(glob.glob(self.output_image_path_ + "/" + '*csv')) is True:

            WST_Area_median = []
            WST_Area_max = []
            WST_Area_min = []
            WST_Area_mean = []
            WST_Area_var = []
            WST_Area_std = []
            WST_Number = []
            WST_Area_Ratio = []
            WST_Orientation = []
            WST_Orientation_var = []
            WST_Density = []
            WST_Area = []
            Filename = []

            # Using for loop to go through all csv files
            if bool(glob.glob(self.output_image_path_ + '/' + 'Statistics.csv')) is True:
                os.remove(glob.glob(self.output_image_path_ + '/' + 'Statistics.csv')[0])
            else:
                self.img_num =1
                for file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                    
                    single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable
                    file_name = file_path[len(self.output_image_path_) + 1:-4]  # Extract the file name from the file path
                    Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
                    
                    Filename.append(",".join([str(item) for item in Split_name]))

                    # Extract the parameters from each single csv files

                    Orientation = single_csv_file["Orientation"]  # Extract the Orientation
                    Whole_stomata_number = single_csv_file["Num_of_whole_stomata"]  # Extract the number of whole stomata
                    Whole_stomata_areas = single_csv_file[single_csv_file['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
                    Whole_stomata_area_ratio = single_csv_file['Whole_stomata_area_ratio']
                    Whole_stomata_density = single_csv_file["Whole_stomata_density"]

                    # Remove the outliers before calculating orientation variance
                    Q1 = np.percentile(Orientation, 25, method = 'midpoint')
                    Q3 = np.percentile(Orientation, 75,method = 'midpoint')

                    IQR = Q3 - Q1

                    upper = np.where(Orientation >= (Q3+1.5*IQR))
                    lower = np.where(Orientation <= (Q1-1.5*IQR))

                    Orientation.drop(upper[0], inplace = True)
                    Orientation.drop(lower[0], inplace = True)

                    # calculate the median, mean, variance of each leaf stomata number, stomata area

                    Whole_stomata_areas_median = np.median(Whole_stomata_areas)
                    Whole_stomata_areas_max = np.max(Whole_stomata_areas)
                    Whole_stomata_areas_min = np.min(Whole_stomata_areas)
                    Whole_stomata_areas_mean = np.mean(Whole_stomata_areas)
                    Whole_stomata_areas_var = np.var(Whole_stomata_areas)
                    Whole_stomata_areas_std = np.std(Whole_stomata_areas)

                    WST_Area_median.append(Whole_stomata_areas_median)
                    WST_Area_max.append(Whole_stomata_areas_max)
                    WST_Area_min.append(Whole_stomata_areas_min)
                    WST_Area_mean.append(Whole_stomata_areas_mean)
                    WST_Area_var.append(Whole_stomata_areas_var)
                    WST_Area_std.append(Whole_stomata_areas_std)
                    WST_Orientation.append(np.median(Orientation))
                    WST_Orientation_var.append(np.var(Orientation))

                    # Extract the number of stomata
                    WST_Density.append(np.mean(Whole_stomata_density))
                    WST_Number.append(np.mean(Whole_stomata_number))
                    WST_Area.append(np.mean(Whole_stomata_areas))
                    WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))
                    
                    self.progressBar.setValue(int((self.img_num / len(glob.glob(self.output_image_path_ + '/' + '*.csv'))*100)))
                    QtWidgets.QApplication.processEvents()

                    self.img_num +=1                    

            # Convert all lists into pd.series

            Filename = pd.Series(Filename, dtype=pd.StringDtype(), name="Filename")
            WST_Number = pd.Series(WST_Number, dtype=pd.Float64Dtype(), name="WST_Number")
            WST_Area_Ratio = pd.Series(WST_Area_Ratio, dtype=pd.Float64Dtype(), name="WST_Area_Ratio")
            WST_Area_median = pd.Series(WST_Area_median, dtype=pd.Float64Dtype(), name="WST_Area_median")
            WST_Area_max = pd.Series(WST_Area_max, dtype=pd.Float64Dtype(), name="WST_Area_max")
            WST_Area_min = pd.Series(WST_Area_min, dtype=pd.Float64Dtype(), name="WST_Area_min")
            WST_Area_mean = pd.Series(WST_Area_mean, dtype=pd.Float64Dtype(), name="WST_Area_mean")
            WST_Area_var = pd.Series(WST_Area_var, dtype=pd.Float64Dtype(), name="WST_Area_var")
            WST_Area_std = pd.Series(WST_Area_std, dtype=pd.Float64Dtype(), name="WST_Area_std")
            WST_Density = pd.Series(WST_Density, dtype=pd.Float64Dtype(), name ="WST_Density")
            WST_Orientation = pd.Series(WST_Orientation, dtype=pd.Float64Dtype(), name ="WST_Orientation")
            WST_Orientation_var = pd.Series(WST_Orientation_var, dtype=pd.Float64Dtype(), name ="WST_Orientation_var")

            # Put all extracted parameters into a data frame
            Output_data = pd.concat(
                [Filename, WST_Number, WST_Area_Ratio,WST_Density, WST_Area_median, WST_Area_max, WST_Area_min,
                WST_Area_mean, WST_Area_var, WST_Area_std, WST_Orientation, WST_Orientation_var], axis=1)
            Output_data = pd.DataFrame(data=Output_data)
            random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
            Output_data.to_excel(self.output_image_path_ + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
            self.show_info_messagebox_group_analysis()

        else:
            self.show_info_messagebox_group_analysis_2()
            pass          


    # #### Put a button to aqure the filedialog

    def loadInputFolder(self):
        self.Inputfilefolder=str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Input Directory"))
        Input_path = self.lineEdit.setText(self.Inputfilefolder)


    def loadOutputFolder(self):
        self.Outputfilefolder = str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Output Directory"))

        Output_path = self.lineEdit_2.setText(self.Outputfilefolder)


    def loadCsv(self):
        self.model = QtGui.QStandardItemModel(StoManager1)
        self.tableView.setModel(self.model)
        self.output_image_path_ = self.lineEdit_2.text()
        if bool(glob.glob(self.output_image_path_ + "/" + '*csv')) is True:
            for file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                fileName = file_path
                with open(fileName, "r") as fileInput:
                    for row in csv.reader(fileInput):    
                        items = [
                            QtGui.QStandardItem(field)
                            for field in row
                        ]
                        self.model.appendRow(items)
        else:
            self.show_info_messagebox_group_analysis_3()
            pass                
                                    
    def clear_model(self):
        # clear the contents of the table view
        for i in range(self.model.rowCount()):
            for j in range(self.model.columnCount()):
                self.model.setData(self.model.index(i, j), "")
                

    def loadExl(self):
        self.model = QtGui.QStandardItemModel(StoManager1)
        self.tableView.setModel(self.model)
        self.output_image_path_ = self.lineEdit_2.text()
        if bool(glob.glob(self.output_image_path_ + "/" + '*csv')) is True:        
            if bool(glob.glob(self.output_image_path_ + '/' + '*.xlsx')):
                for file_path in glob.glob(self.output_image_path_ + '/' + '*.xlsx'):
                    newpath =  self.output_image_path_ + f'/New_folder' 
                    if not os.path.exists(newpath):
                        os.makedirs(newpath)
                        self.folder = newpath
                    else:
                        self.folder = newpath
                    for filenames in os.listdir(self.folder):
                        file_paths = os.path.join(self.folder, filenames)
                        try:
                            if os.path.isfile(file_paths) or os.path.islink(file_paths):
                                os.unlink(file_paths)
                            elif os.path.isdir(file_paths):
                                shutil.rmtree(file_paths)
                        except Exception as e:
                            print('Failed to delete %s. Reason: %s' % (file_paths, e))   
                    pd.read_excel(file_path).to_csv(newpath + '/' + ''.join(random.choices(string.ascii_uppercase, k=4)) +".csv", 
                    index = None,
                    header=True)
                    for fileName in glob.glob(newpath + '/' + '*.csv'):
                        with open(fileName, "r") as fileInput:
                            for row in csv.reader(fileInput):    
                                items = [
                                    QtGui.QStandardItem(field)
                                    for field in row
                                ]
                                self.model.appendRow(items)
            else:
                self.show_info_messagebox_group_analysis_2()        

        else:
            
            self.show_info_messagebox_question()
            pass

    def show_info_messagebox(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Yeah! Process finished!")
        
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Processor 🐸")
        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Now you can check your output or do statistical analysis. 🐻")
        
        # start the app
        x = msg.exec_()

    def popup_button(self, i):
        print(i.text())

    def show_info_messagebox_question(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps...  there is no Statistics file. 🦊")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Statistical results preview 📊")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Have you done Group or Non-group analysis? 👀")
        
        # start the app
        x= msg.exec_()
                    
    def show_info_messagebox_group_analysis(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Sweet 😘~  Now you can preview statistical results. 🐮")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Statistical analysis 🐼")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("💕")
        
        # start the app
        x= msg.exec_()  

    def show_info_messagebox_group_analysis1(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps... Now this function is only applicable for our Populus dataset.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Group analysis 📈")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("I will add a customizable function for our users very soon 🐌.")
        
        # start the app
        x= msg.exec_() 

    def show_info_messagebox_group_analysis_2(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps... It looks like you haven't done processing images.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Statistical analysis 📈")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define file paths, start process, and give it one more try  🐌.")
        
        # start the app
        x= msg.exec_() 

    def show_info_messagebox_group_analysis_3(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps... It looks like you haven't done processing images.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Preview exported results 📈")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define file paths, start process, and give it one more try  🐌.")
        
        # start the app
        x= msg.exec_() 
                
    def show_info_messagebox_normalize(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Please note that now filename normalizer is only applicable for our Populus dataset. It will work if you are playing our Populus dataset.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 FileName Normalizer 📶")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("I will add a customizable function for our users very soon 💕.")
        
        # start the app
        x= msg.exec_() 

    def show_info_messagebox_normalize_2(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps... It looks like you haven't define input file path.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 FileName Normalizer 📶")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define and input path and try one more time 💕.")
        
        # start the app
        x= msg.exec_() 

    def show_info_messagebox_original_image(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Heads-up 🦉 You can scroll horizontally or vertically to view the whole image (●’◡’●).")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 View images 👨‍🏫")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("I will add a customizable function for our users very soon 💕.")
        
        # start the app
        x= msg.exec_() 

    def messagebox_define_input_path(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Oh... You may forget define input path....")
        
        # setting Message box window title
        msg.setWindowTitle("StoManager1 View Original images 🐸")
        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define input image path, and try one more time. 🐻")
        
        # start the app
        x = msg.exec_()

    def messagebox_define_input_path_2(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Oh... You may forget define input path....")
        
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Processor images 🐸")
        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define input image path, and try one more time. 🐻")
        
        # start the app
        x = msg.exec_()

    def messagebox_define_output_path(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)

        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Oh... You may forget define output path....")
        
        # setting Message box window title
        msg.setWindowTitle("StoManager1 View Detected images 🐸")
        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)

        msg.setInformativeText("Please define output image path, and try one more time. 🐻")
        
        # start the app
        x = msg.exec_()

# webbrowser

    def web_link_guthub(self):
        webbrowser.open('https://github.com/JiaxinWang123/StoManager.git')

    def web_link_google_scholar(self):
        webbrowser.open("https://scholar.google.com/citations?user=7be6E64AAAAJ&hl=en")

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    StoManager1 = QtWidgets.QMainWindow()
    ui = Ui_StoManager1()
    ui.setupUi(StoManager1)
    StoManager1.show()
    sys.exit(app.exec_())
