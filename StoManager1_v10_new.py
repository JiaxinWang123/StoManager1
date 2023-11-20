#### Imports ####
from tkinter import *
import cv2
import numpy as np
import glob
import random
import pandas as pd
import math
import string
import re
from PyQt5 import QtCore, QtGui, QtWidgets
import csv
from PyQt5.QtGui import QPixmap
import webbrowser
#########################

#### Import packages for segment models####
from ultralytics import YOLO
from shapely.geometry import Polygon
from shapely import Point
import os
import torch
import os, shutil
import glob
from scipy.spatial import distance, distance_matrix
from scipy.sparse.csgraph import minimum_spanning_tree
########################

# Another Window #
from PyQt5.QtWidgets import QMainWindow
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from qtpy.QtCore import QThread, Signal
import subprocess
import io


class Ui_StoManager1_training(QMainWindow):
    def __init__(self):
        super(Ui_StoManager1_training, self).__init__()
        self.setupUi(self)
        self.training_runner = None  # Store the training thread
        self.training_process = None  # Store the training subprocess
        self.training_running = False

    def setupUi(self, StoManager1):

        StoManager1.setObjectName("StoManager1")
        StoManager1.resize(1200, 800)
        StoManager1.setWindowIcon(QtGui.QIcon("StoManager.ico"))
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
"background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:1 rgba(29, 141, 162, 255));")
        self.centralwidget = QtWidgets.QWidget(StoManager1)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalWidget_2 = QtWidgets.QWidget(self.centralwidget)
        self.verticalWidget_2.setObjectName("verticalWidget_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.verticalWidget_2)
        self.verticalLayout_3.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.verticalWidget_2)
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
        self.pushButton_8 = QtWidgets.QPushButton(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setStyleSheet("background-color: rgb(193, 101, 68);\n"
"")
        self.pushButton_8.setObjectName("pushButton_8")
        self.pushButton_8.clicked.connect(self.loadInputFolder)
        self.horizontalLayout.addWidget(self.pushButton_8)
        self.verticalLayout_3.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.verticalWidget_2)
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
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.pushButton_9 = QtWidgets.QPushButton(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setStyleSheet("background-color: rgb(193, 101, 68);")
        self.pushButton_9.setObjectName("pushButton_9")
        self.pushButton_9.clicked.connect(self.loadOutputFolder)
        self.horizontalLayout_2.addWidget(self.pushButton_9)
        self.verticalLayout_3.addLayout(self.horizontalLayout_2)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.verticalLayout_2.addLayout(self.horizontalLayout_9)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.verticalWidget_2)
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
        self.verticalLayout.addWidget(self.lineEdit_4)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.verticalWidget_2)
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
        self.verticalLayout.addWidget(self.lineEdit_5)
        self.lineEdit_6 = QtWidgets.QLineEdit(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_6.setFont(font)
        self.lineEdit_6.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_6.setClearButtonEnabled(True)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.verticalLayout.addWidget(self.lineEdit_6)
        self.lineEdit_7 = QtWidgets.QLineEdit(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_7.setFont(font)
        self.lineEdit_7.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_7.setClearButtonEnabled(True)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.verticalLayout.addWidget(self.lineEdit_7)
        self.lineEdit_9 = QtWidgets.QLineEdit(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.lineEdit_9.setFont(font)
        self.lineEdit_9.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_9.setClearButtonEnabled(True)
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.verticalLayout.addWidget(self.lineEdit_9)
        self.checkBox = QtWidgets.QCheckBox(self.verticalWidget_2)
        self.checkBox.setStyleSheet("color : rgb(255, 255, 255)")
        self.checkBox.setChecked(True)
        self.checkBox.setObjectName("checkBox")
        self.verticalLayout.addWidget(self.checkBox)
        self.checkBox_2 = QtWidgets.QCheckBox(self.verticalWidget_2)
        self.checkBox_2.setStyleSheet("color : rgb(255, 255, 255)")
        self.checkBox_2.setObjectName("checkBox_2")
        self.verticalLayout.addWidget(self.checkBox_2)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.pushButton_3 = QtWidgets.QPushButton(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(91, 215, 244);\n"
"")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.Check_input_path_folder)
        self.horizontalLayout_3.addWidget(self.pushButton_3)
        self.pushButton_6 = QtWidgets.QPushButton(self.verticalWidget_2)
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
        self.pushButton_6.clicked.connect(self.stop_script)
        self.horizontalLayout_3.addWidget(self.pushButton_6)
        self.pushButton_7 = QtWidgets.QPushButton(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setStyleSheet("background-color: rgb(15, 160, 123);")
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_7.clicked.connect(self.openFileExplorer)
        self.horizontalLayout_3.addWidget(self.pushButton_7)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        # self.progressBar = QtWidgets.QProgressBar(self.verticalWidget_2)
        # self.progressBar.setProperty("value", 0)
        # self.progressBar.setObjectName("progressBar")
        # self.verticalLayout_2.addWidget(self.progressBar)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.verticalWidget_2)
        self.plainTextEdit.setAutoFillBackground(False)
        self.plainTextEdit.setStyleSheet("background-color: rgb(30, 30, 30); color: white\n"
"\n"
"")
        self.plainTextEdit.setPlainText("")
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout_2.addWidget(self.plainTextEdit)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_3 = QtWidgets.QLabel(self.verticalWidget_2)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(214, 214, 214);")
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_10.addWidget(self.label_3)
        self.label_9 = QtWidgets.QLabel(self.verticalWidget_2)
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
        self.verticalLayout_2.addLayout(self.horizontalLayout_10)
        self.verticalLayout_3.addLayout(self.verticalLayout_2)
        self.gridLayout.addWidget(self.verticalWidget_2, 0, 0, 1, 1)
        StoManager1.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(StoManager1)
        self.statusbar.setObjectName("statusbar")
        StoManager1.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(StoManager1)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1221, 22))
        self.menubar.setObjectName("menubar")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuTraining = QtWidgets.QMenu(self.menubar)
        self.menuTraining.setObjectName("menuTraining")
        StoManager1.setMenuBar(self.menubar)
        self.actionH = QtWidgets.QAction(StoManager1)
        self.actionH.setObjectName("actionH")
        self.action = QtWidgets.QAction(StoManager1)
        self.action.setObjectName("action")
        self.actionYOLOv8_seg_x = QtWidgets.QAction(StoManager1)
        self.actionYOLOv8_seg_x.setObjectName("actionYOLOv8_seg_x")
        self.actionGoogle_scholar = QtWidgets.QAction(StoManager1)
        self.actionGoogle_scholar.setObjectName("actionGoogle_scholar")
        self.actionarXiv = QtWidgets.QAction(StoManager1)
        self.actionarXiv.setObjectName("actionarXiv")
        self.actionGitHub = QtWidgets.QAction(StoManager1)
        self.actionGitHub.setObjectName("actionGitHub")
        self.actionMy_Homepage = QtWidgets.QAction(StoManager1)
        self.actionMy_Homepage.setObjectName("actionMy_Homepage")
        self.actionMy_Homepage.triggered.connect(self.web_link_ultralytics)
        self.actionEarly_versions = QtWidgets.QAction(StoManager1)
        self.actionEarly_versions.setObjectName("actionEarly_versions")
        self.actionOnly_whole_stomata = QtWidgets.QAction(StoManager1)
        self.actionOnly_whole_stomata.setCheckable(True)
        self.actionOnly_whole_stomata.setObjectName("actionOnly_whole_stomata")
        self.actionOnly_stomata_aperture = QtWidgets.QAction(StoManager1)
        self.actionOnly_stomata_aperture.setCheckable(True)
        self.actionOnly_stomata_aperture.setObjectName("actionOnly_stomata_aperture")
        self.actionOnly_guard_cell = QtWidgets.QAction(StoManager1)
        self.actionOnly_guard_cell.setCheckable(True)
        self.actionOnly_guard_cell.setObjectName("actionOnly_guard_cell")
        self.actionAll_metrics = QtWidgets.QAction(StoManager1)
        self.actionAll_metrics.setCheckable(True)
        self.actionAll_metrics.setChecked(True)
        self.actionAll_metrics.setObjectName("actionAll_metrics")
        self.actionGroup_Analysis = QtWidgets.QAction(StoManager1)
        self.actionGroup_Analysis.setObjectName("actionGroup_Analysis")
        self.actionOpen_training_window = QtWidgets.QAction(StoManager1)
        self.actionOpen_training_window.setObjectName("actionOpen_training_window")
        self.actionOpen_training_window.triggered.connect(self.web_link_YOLOv8)
        self.menuHelp.addAction(self.actionMy_Homepage)
        self.menuTraining.addAction(self.actionOpen_training_window)
        self.menubar.addAction(self.menuTraining.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.checkBox.stateChanged.connect(self.update_amp_state)
        self.checkBox_2.stateChanged.connect(self.update_device_state)
        self.retranslateUi(StoManager1)
        QtCore.QMetaObject.connectSlotsByName(StoManager1)
    
    def update_amp_state(self):
        self.amp = "True" if self.checkBox.isChecked() else "False"

    def update_device_state(self):
        self.device = "cpu" if self.checkBox_2.isChecked() else "gpu"

    def retranslateUi(self, StoManager1):
        _translate = QtCore.QCoreApplication.translate
        StoManager1.setWindowTitle(_translate("StoManager1", "StoManager1_model_trainer"))
        self.lineEdit.setText(_translate("StoManager1", "Select training data input data.yaml"))
        self.pushButton_8.setText(_translate("StoManager1", "Input"))
        self.lineEdit_2.setText(_translate("StoManager1", "Select model_training_in_app.exe file"))
        self.pushButton_9.setText(_translate("StoManager1", "Trainer"))
        self.lineEdit_4.setText(_translate("StoManager1", "Number of training epochs, default is 1000"))
        self.lineEdit_5.setText(_translate("StoManager1", "Image size, default is 640"))
        self.lineEdit_6.setText(_translate("StoManager1", "Batch, 2-64 depends on your GPU or CPU, default is 2"))
        self.lineEdit_7.setText(_translate("StoManager1", "Fliplr, default is 0"))
        self.lineEdit_9.setText(_translate("StoManager1", "Workers, default is 0"))
        self.checkBox.setText(_translate("StoManager1", "AMP Check, uncheck it if you got nan for training metrics"))
        self.checkBox_2.setText(_translate("StoManager1", "Train on CPU, check it if you don't have a powerful GPU"))        
        self.pushButton_3.setText(_translate("StoManager1", "Start Training"))
        self.pushButton_6.setText(_translate("StoManager1", "Stop Training"))
        self.pushButton_7.setText(_translate("StoManager1", "Check training result"))
        self.label_3.setText(_translate("StoManager1", "© Jiaxin Wang.  For questions or requests 📧  coolwjx@foxmail.com; jw3994@msstate.edu"))
        self.label_9.setText(_translate("StoManager1", "💕   LU"))
        self.menuHelp.setTitle(_translate("StoManager1", "Help"))
        self.menuTraining.setTitle(_translate("StoManager1", "Training"))
        self.actionH.setText(_translate("StoManager1", "H"))
        self.action.setText(_translate("StoManager1", ")"))
        self.actionYOLOv8_seg_x.setText(_translate("StoManager1", "YOLOv8-seg-x"))
        self.actionGoogle_scholar.setText(_translate("StoManager1", "Google scholar"))
        self.actionarXiv.setText(_translate("StoManager1", "arXiv"))
        self.actionGitHub.setText(_translate("StoManager1", "GitHub"))
        self.actionMy_Homepage.setText(_translate("StoManager1", "Ultralytics"))
        self.actionEarly_versions.setText(_translate("StoManager1", "Early versions"))
        self.actionOnly_whole_stomata.setText(_translate("StoManager1", "Only whole_stomata"))
        self.actionOnly_stomata_aperture.setText(_translate("StoManager1", "Only stomata (aperture)"))
        self.actionOnly_guard_cell.setText(_translate("StoManager1", "Only guard cell"))
        self.actionAll_metrics.setText(_translate("StoManager1", "All metrics"))
        self.actionGroup_Analysis.setText(_translate("StoManager1", "Group Analysis"))
        self.actionOpen_training_window.setText(_translate("StoManager1", "Train YOLOv8-seg-x"))

        self.training_runner = TrainingRunner()
        self.training_runner.update_signal.connect(self.update_console)
        self.training_running = False

    def Check_input_path_folder(self):
        data_path = self.lineEdit.text()
        model_path = self.lineEdit_2.text()

        if data_path and model_path:
            # Both paths are provided, you can start training here
            if data_path == "Select training data input data.yaml" or model_path == "Select model_training_in_app.exe file" or data_path == "" or model_path == "":
                self.show_message()
            else:
                self.start_training()
        else:
            # Show a message box indicating that both paths are required
            self.show_message("Error", "Please select both data.yaml and trainer files.")

    def show_message(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Oh... You may forget define path of training data and model trainer....")        
        # setting Message box window title
        msg.setWindowTitle("Define your data file and model trainer 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please define both file path, and try one more time. 🐻")       
        # start the app
        msg.exec_()

    def show_message_2(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Oh... You haven't started training yet....")        
        # setting Message box window title
        msg.setWindowTitle("Define your data file and model trainer and start your training 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please define your data file and model trainer and start your training, and try one more time. 🐻")       
        # start the app
        msg.exec_()        

    def openFileExplorer(self):
        # os.chdir(os.path.dirname(sys.executable))  # Comment this line out if you use it in python interpreter

        # Append the "runs/segment" subdirectory
        desired_path = os.path.join(os.getcwd(), "runs", "segment")
        if os.path.exists(desired_path) and os.path.isdir(desired_path):
            options = QtWidgets.QFileDialog.Options()
            QtWidgets.QFileDialog.getOpenFileName(self, "Open Directory", desired_path, "All Files (*);;", options=options)
            self.pushButton_7.setEnabled(True)  # Enable the "Check training result" button
        else:
            self.show_message_2()
            self.pushButton_7.setEnabled(False)  # Disable the "Check training result" button

    def web_link_ultralytics(self):
        """ """
        webbrowser.open("https://github.com/ultralytics/ultralytics")

    def web_link_YOLOv8(self):
        """ """
        webbrowser.open("https://docs.ultralytics.com/modes/train/#introduction")

    def loadInputFolder(self):
        """ Load Input file folder"""
        options = QtWidgets.QFileDialog.Options()
        self.Inputfilefolder, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select data.yaml file", "", "Yaml Files (*.yaml);;All Files (*)", options=options)
        Input_path = self.lineEdit.setText(self.Inputfilefolder)

    def loadOutputFolder(self):
        """ Load Output folder"""
        options = QtWidgets.QFileDialog.Options()
        self.Outputfilefolder, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select model_trainer.exe file", "", "Excutable Files (*.exe);;All Files (*)", options=options)
        Output_path = self.lineEdit_2.setText(self.Outputfilefolder)

    def get_parameters(self):
        # Retrieve parameters from the QLineEdit widgets
        script_path = self.lineEdit_2.text()
        data_input_path = self.lineEdit.text()
        epoch = self.lineEdit_4.text()
        imgsz = self.lineEdit_5.text()
        batch = self.lineEdit_6.text()
        fliplr = self.lineEdit_7.text()
        workers = self.lineEdit_9.text()
        weights_path = self.lineEdit_2.text()
        amp = "True" if self.checkBox.isChecked() else "False"
        device = "cpu" if self.checkBox_2.isChecked() else "gpu"

        return script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path, amp, device

    def stop_script(self):
        if self.training_running:
            if self.training_runner and self.training_runner.isRunning():
                self.training_runner.terminate()
                self.training_runner.wait()
            if self.training_process and self.training_process.poll() is None:
                self.training_process.terminate()
                self.training_process.wait()
            self.training_running = False
            self.pushButton_3.setEnabled(True)  # Enable the "Start Training" button

    def start_training(self):
        if not self.training_running:
            self.training_running = True
            self.pushButton_3.setEnabled(False)
            self.plainTextEdit.clear()

            # Get parameters from QLineEdit widgets
            script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path, amp, device = self.get_parameters()

            weights_path = script_path  # Specify the desired file path here
            folder_path = os.path.dirname(weights_path)
            weights_path = os.path.join(folder_path, 'yolov8x-seg.pt')

            if epoch == "Number of training epochs, default is 1000" or epoch == "" or epoch == " ":
                epoch = 1000
            else:
                epoch = self.lineEdit_4.text()

            if imgsz == "Image size, default is 640" or imgsz == "" or imgsz == " ":
                imgsz = 640
            else:
                imgsz = self.lineEdit_5.text()

            if batch == "Batch, 2-64 depends on your GPU or CPU, default is 2" or batch == "" or batch == " ":
                batch = 2
            else:
                batch = self.lineEdit_6.text()

            if fliplr == "Fliplr, default is 0" or fliplr == "" or fliplr == " ":
                fliplr = 0
            else:
                fliplr = self.lineEdit_7.text()

            if workers == "Workers, default is 0" or workers == "" or workers == " ":
                workers = 0
            else:
                workers = self.lineEdit_9.text()

            # Set parameters for the TrainingRunner object
            self.training_runner = TrainingRunner()
            self.training_runner.set_parameters(script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path,amp,device)
            self.training_runner.update_signal.connect(self.update_console)
            self.training_runner.finished_signal.connect(self.stop_script)  # Connect to the stop_script method

            # Start the training thread
            self.training_runner.start()

    def update_console(self, text):
        self.plainTextEdit.appendPlainText(text)

class TrainingRunner(QThread):
    update_signal = Signal(str)
    finished_signal = Signal()

    def __init__(self):
        super(TrainingRunner, self).__init__()
        self.script_path = None
        self.data_input_path = None
        self.epoch = None
        self.imgsz = None
        self.batch = None
        self.fliplr = None
        self.workers = None
        self.weights_path = None
        self.amp = None
        self.device = None

    def set_parameters(self, script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path, amp,device):
        self.script_path = script_path
        self.model_output_path = self.script_path
        self.data_input_path = data_input_path
        self.epoch = epoch
        self.imgsz = imgsz
        self.batch = batch
        self.fliplr = fliplr
        self.workers = workers
        self.weights_path = weights_path
        self.amp = amp
        self.device = device

    def run(self):
        try:
            process = subprocess.Popen(
                ['python',self.script_path, self.data_input_path, self.model_output_path, str(self.epoch), str(self.imgsz),
                 str(self.batch), str(self.fliplr), str(self.workers), self.weights_path, self.amp, self.device],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.PIPE,
                text=True,  # Use text instead of universal_newlines
                close_fds=True,
                env=os.environ
            )

            # Create a non-blocking file object for stdout
            stdout_reader = io.open(process.stdout.fileno(), 'r', encoding='utf-8', errors='replace')

            while process.poll() is None:  # Check if the subprocess is still running
                line = stdout_reader.readline()
                if not line:
                    break
                self.update_signal.emit(line.strip())

            # Emit the signal when the subprocess is finished
            self.finished_signal.emit()
        except Exception as e:
            self.update_signal.emit(f"Error: {str(e)}")
            print(f"Exception during run: {e}")
        finally:
            if process.poll() is None:  # Check if the subprocess is still running
                process.terminate()  # Terminate the subprocess if it's still running
                process.wait()  # Wait for the subprocess to finish
    def update_console(self, text):
        self.plainTextEdit.append(text)

#### Main Window ####
class Ui_StoManager1(object):
    def __init__(self):
        super(Ui_StoManager1, self).__init__()
        self.setupUi(StoManager1)
        self.selected_image_index = 0
        self.img_num = 1
        self.folder =str()
        self.p = 465

    def setupUi(self, StoManager1):
        StoManager1.setObjectName("StoManager1_v10_Seg-x_Hardwoods")
        StoManager1.resize(1200, 800)
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
"background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:1 rgba(29, 141, 162, 255));")
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
        self.pushButton_2.clicked.connect(self.Output_previous_img) 
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
        self.pushButton_5.clicked.connect(lambda:[self.Check_Output_path(), self.show_info_messagebox_original_image()]) 
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
        self.pushButton_11.clicked.connect(self.Output_next_img)  
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
        self.gridLayout.addLayout(self.gridLayout_2, 0, 2, 1, 1)
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
        self.pushButton_8.setStyleSheet("background-color: rgb(196, 124, 180);\n"
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
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setStyleSheet("background-color: rgb(196, 124, 180);")
        self.pushButton_9.setObjectName("pushButton_9")
        self.pushButton_9.clicked.connect(self.loadOutputFolder)
        self.horizontalLayout_2.addWidget(self.pushButton_9)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.YOLOv8_seg_x = QtWidgets.QCheckBox(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.YOLOv8_seg_x.setFont(font)
        self.YOLOv8_seg_x.setAutoFillBackground(False)
        self.YOLOv8_seg_x.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.YOLOv8_seg_x.setChecked(True)
        self.YOLOv8_seg_x.setTristate(False)
        self.YOLOv8_seg_x.setObjectName("YOLOv8_seg_x")
        self.horizontalLayout_5.addWidget(self.YOLOv8_seg_x)
        self.verticalLayout.addLayout(self.horizontalLayout_5)
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
        self.verticalLayout.addWidget(self.lineEdit_4)
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
        self.pushButton_3.setStyleSheet("background-color: rgb(79, 186, 112);\n"
"")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.check_input_Path_run) 
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
        self.pushButton_7.setStyleSheet("background-color: rgb(189, 214, 67);")
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_7.clicked.connect(self.Statistics) 
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
        self.pushButton_15.clicked.connect(lambda: [self.clear_model(), self.Load_csv()]) 
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
        self.pushButton_16.clicked.connect(lambda: [self.clear_model(), self.Load_Elx()]) 
        self.horizontalLayout_9.addWidget(self.pushButton_16)
        self.verticalLayout_2.addLayout(self.horizontalLayout_9)
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.tableView.setFont(font)
        self.tableView.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"\n"
"")
        self.tableView.setObjectName("tableView")
        self.verticalLayout_2.addWidget(self.tableView)
        self.verticalLayout_4.addLayout(self.verticalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout_4, 0, 0, 1, 1)
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
        self.statusbar = QtWidgets.QStatusBar(StoManager1)
        self.statusbar.setObjectName("statusbar")
        StoManager1.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(StoManager1)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1155, 22))
        self.menubar.setObjectName("menubar")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuAnalysis = QtWidgets.QMenu(self.menubar)
        self.menuAnalysis.setObjectName("menuAnalysis")
        self.menuTraining = QtWidgets.QMenu(self.menubar)
        self.menuTraining.setObjectName("menuTraining")
        StoManager1.setMenuBar(self.menubar)
        self.actionH = QtWidgets.QAction(StoManager1)
        self.actionH.setObjectName("actionH")
        self.action = QtWidgets.QAction(StoManager1)
        self.action.setObjectName("action")
        self.actionYOLOv8_seg_x = QtWidgets.QAction(StoManager1)
        self.actionYOLOv8_seg_x.setObjectName("actionYOLOv8_seg_x")
        self.actionGoogle_scholar = QtWidgets.QAction(StoManager1)
        self.actionGoogle_scholar.setObjectName("actionGoogle_scholar")
        self.actionGoogle_scholar.triggered.connect(self.web_link_google_scholar)
        self.actionarXiv = QtWidgets.QAction(StoManager1)
        self.actionarXiv.setObjectName("actionarXiv")
        self.actionarXiv.triggered.connect(self.web_link_arxiv)
        self.actionGitHub = QtWidgets.QAction(StoManager1)
        self.actionGitHub.setObjectName("actionGitHub")
        self.actionGitHub.triggered.connect(self.web_link_github)
        self.actionMy_Homepage = QtWidgets.QAction(StoManager1)
        self.actionMy_Homepage.setObjectName("actionMy_Homepage")
        self.actionMy_Homepage.triggered.connect(self.web_link_Homepage) 
        self.actionEarly_versions = QtWidgets.QAction(StoManager1)
        self.actionEarly_versions.setObjectName("actionEarly_versions")
        self.actionEarly_versions.triggered.connect(self.web_link_Early_versions) 
        self.actionOnly_whole_stomata = QtWidgets.QAction(StoManager1)
        self.actionOnly_whole_stomata.setCheckable(True)
        self.actionOnly_whole_stomata.setObjectName("actionOnly_whole_stomata")
        self.actionOnly_stomata_aperture = QtWidgets.QAction(StoManager1)
        self.actionOnly_stomata_aperture.setCheckable(True)
        self.actionOnly_stomata_aperture.setObjectName("actionOnly_stomata_aperture")
        self.actionOnly_guard_cell = QtWidgets.QAction(StoManager1)
        self.actionOnly_guard_cell.setCheckable(True)
        self.actionOnly_guard_cell.setObjectName("actionOnly_guard_cell")
        self.actionAll_metrics = QtWidgets.QAction(StoManager1)
        self.actionAll_metrics.setCheckable(True)
        self.actionAll_metrics.setChecked(True)
        self.actionAll_metrics.setObjectName("actionAll_metrics")
        self.actionGroup_Analysis = QtWidgets.QAction(StoManager1)
        self.actionGroup_Analysis.setCheckable(True)
        self.actionGroup_Analysis.setObjectName("actionGroup_Analysis")
        self.actionGroup_Analysis.triggered.connect(self.StatisticsGroup)
        self.actionOpen_training_window = QtWidgets.QAction(StoManager1)
        self.actionOpen_training_window.setObjectName("actionOpen_training_window")
        self.actionOpen_training_window.triggered.connect(self.open_training_window)
        self.menuHelp.addAction(self.actionGoogle_scholar)
        self.menuHelp.addAction(self.actionarXiv)
        self.menuHelp.addAction(self.actionGitHub)
        self.menuHelp.addAction(self.actionMy_Homepage)
        self.menuHelp.addSeparator()
        self.menuHelp.addAction(self.actionEarly_versions)
        self.menuAnalysis.addAction(self.actionOnly_whole_stomata)
        self.menuAnalysis.addAction(self.actionOnly_stomata_aperture)
        self.menuAnalysis.addAction(self.actionOnly_guard_cell)
        self.menuAnalysis.addSeparator()
        self.menuAnalysis.addAction(self.actionAll_metrics)
        self.menuAnalysis.addSeparator()
        self.menuAnalysis.addAction(self.actionGroup_Analysis)
        self.menuTraining.addAction(self.actionOpen_training_window)
        self.menubar.addAction(self.menuAnalysis.menuAction())
        self.menubar.addAction(self.menuTraining.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(StoManager1)
        QtCore.QMetaObject.connectSlotsByName(StoManager1)

    def open_training_window(self):
        self.second_window = Ui_StoManager1_training()
        self.second_window.show()

    def retranslateUi(self, StoManager1):
        _translate = QtCore.QCoreApplication.translate
        StoManager1.setWindowTitle(_translate("StoManager1", "StoManager1 v.1.0.0."))
        self.pushButton.setText(_translate("StoManager1", "<< View previous"))
        self.pushButton_4.setText(_translate("StoManager1", "View original images"))
        self.pushButton_10.setText(_translate("StoManager1", "View next >>"))
        self.pushButton_2.setText(_translate("StoManager1", "<< View previous"))
        self.pushButton_5.setText(_translate("StoManager1", "View detected images"))
        self.pushButton_11.setText(_translate("StoManager1", "View next >>"))
        self.lineEdit.setText(_translate("StoManager1", "Input path"))
        self.pushButton_8.setText(_translate("StoManager1", "Input"))
        self.lineEdit_2.setText(_translate("StoManager1", "Output path"))
        self.pushButton_9.setText(_translate("StoManager1", "Output"))
        self.YOLOv8_seg_x.setText(_translate("StoManager1", "Segment model using trained YOLOv8-seg-x"))
        self.lineEdit_4.setText(_translate("StoManager1", "Input pixels in 0.1 mm, the default is: 465"))
        self.lineEdit_5.setText(_translate("StoManager1", "Confidence threshold for detection the default is 0.25"))
        self.pushButton_3.setText(_translate("StoManager1", "Start Process"))
        self.pushButton_6.setText(_translate("StoManager1", "Normalize File"))
        self.pushButton_7.setText(_translate("StoManager1", "Statistical Analysis"))
        self.pushButton_15.setText(_translate("StoManager1", "Preview exported table for single image"))
        self.pushButton_16.setText(_translate("StoManager1", "Preview exported statistics for all images"))
        self.label_3.setText(_translate("StoManager1", "© Jiaxin Wang.  For questions or requests 📧  coolwjx@foxmail.com; jw3994@msstate.edu"))
        self.label_9.setText(_translate("StoManager1", "💕   LU"))
        self.menuHelp.setTitle(_translate("StoManager1", "Help"))
        self.menuAnalysis.setTitle(_translate("StoManager1", "Analysis"))
        self.menuTraining.setTitle(_translate("StoManager1", "Training"))
        self.actionH.setText(_translate("StoManager1", "H"))
        self.action.setText(_translate("StoManager1", ")"))
        self.actionYOLOv8_seg_x.setText(_translate("StoManager1", "YOLOv8-seg-x"))
        self.actionGoogle_scholar.setText(_translate("StoManager1", "Google scholar"))
        self.actionarXiv.setText(_translate("StoManager1", "arXiv"))
        self.actionGitHub.setText(_translate("StoManager1", "GitHub"))
        self.actionMy_Homepage.setText(_translate("StoManager1", "My Homepage"))
        self.actionEarly_versions.setText(_translate("StoManager1", "Early versions"))
        self.actionOnly_whole_stomata.setText(_translate("StoManager1", "Only whole_stomata"))
        self.actionOnly_stomata_aperture.setText(_translate("StoManager1", "Only stomata (aperture)"))
        self.actionOnly_guard_cell.setText(_translate("StoManager1", "Only guard cell"))
        self.actionAll_metrics.setText(_translate("StoManager1", "All metrics"))
        self.actionGroup_Analysis.setText(_translate("StoManager1", "Group Analysis"))
        self.actionOpen_training_window.setText(_translate("StoManager1", "Train YOLOv8-seg-x"))

    def Analysis_S(self):
        """ Do stomatal detection and meauring.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        if self.YOLOv8_seg_x.isChecked():
            self.Check_input_path_folder()
        else:
            self.Box_model()

    def Statistics(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        if self.YOLOv8_seg_x.isChecked():
            self.Stata()
        else:
            self.Stomata_no_groups_analysis()

    def StatisticsGroup(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        if self.YOLOv8_seg_x.isChecked():
            self.Stomata_group_analysis_seg()
        else:
            self.Stomata_group_analysis()

    def Load_csv(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        if self.YOLOv8_seg_x.isChecked():
            self.loadCsv_Segment_model()
        else:
            self.loadCsv_Box_model()

    def Load_Elx(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        if self.YOLOv8_seg_x.isChecked():
            self.loadExl_Segment_model()
        else:
            self.loadExl_Box_model()

    def Check_Output_path(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        try:
            if self.YOLOv8_seg_x.isChecked():
                self.check_output_Path_Segment_model()
            else:
                self.check_output_Path_Box_model()
        except Exception as e:
            self.outofIndex(e)

    def Output_next_img(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        try:
            if self.YOLOv8_seg_x.isChecked():
                self.output_next_image_clicked_Segment_model()
            else:
                self.output_next_image_clicked_Box_model()
        except Exception as e:
            self.outofIndex(e)

    def Output_previous_img(self):
        """ Do statistical analysis of the output results.
        check if the YOLOv8-seg-x function is checked.
        if it is checked, run `Guardcell()` function, which uses YOLOv8-seg-x sementation-based model;
        if it is not checked, run `run_analyze()` function, which uses YOLOv3 bounding box-based model
        """
        try:
            if self.YOLOv8_seg_x.isChecked():
                self.output_previous_image_clicked_Segment_model()
            else:
                self.output_previous_image_clicked_Box_model()
        except Exception as e:
            self.outofIndex(e)

    def check_input_Path(self):
        """ """
        try:
            self.input_image_path_ = self.lineEdit.text()
            # Specify the directory containing the images
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:
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
        except Exception as e:
            self.outofIndex(e)

    def check_input_Path_run(self):
        """ Check if the input path is defined and contains image files. """
        try:
            self.input_image_path_ = self.lineEdit.text()

            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)
            
            if has_image_files:
                self.Analysis_S()
            else:
                self.messagebox_define_input_path()
        except Exception as e:
            self.outofIndex(e)


    def check_output_Path_Box_model(self):
        """ Check if the YOLOv8-seg-x model selected, if not, using box model instead. """
        try:
            self.output_image_path_ = self.lineEdit_2.text()
                    # get first image file and show it
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:
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
        except Exception as e:
            self.outofIndex(e)

    def check_output_Path_Segment_model(self):
        """ """
        try:
            self.output_image_path_ = self.lineEdit_2.text()

            output_image_path_ = glob.glob(os.path.join(str(self.output_image_path_), "Predict_output/Output_csv", "*"))

            image_files_2 = []
            for folder in output_image_path_:
                for f in glob.glob(folder+'/*.jpg'):
                    image_files_2.append(f)

            if bool(output_image_path_) is True:
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
        except Exception as e:
            self.outofIndex(e)

# next image

    def input_next_image_clicked(self):
        """ """
        # get first image file and show it
        try:
            self.input_image_path_ = self.lineEdit.text()
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:

                if self.selected_image_index in range(len(image_files_2)-1):
                    # If reached the last image, set index to start over
                    if self.selected_image_index == len(image_files_2) - 1:
                        self.selected_image_index = 0
                    else:
                        self.selected_image_index += 1
                else:
                    self.selected_image_index -= 1
                scene = QtWidgets.QGraphicsScene(StoManager1)
                pixmap = QPixmap(image_files_2[self.selected_image_index])
                item = QtWidgets.QGraphicsPixmapItem(pixmap)
                scene.addItem(item)
                self.graphicsView.setScene(scene)
                #self.selected_image_index += 1
            else:
                self.messagebox_define_input_path()
                pass

        except Exception as e:
            self.outofIndex(e)

#next labeled image
    def output_next_image_clicked_Box_model(self):
        """ """
        # get first image file and show it
        try:
            self.output_image_path_ = self.lineEdit_2.text()
            image_files = glob.glob(self.output_image_path_ + "/" + '*jpg')
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:

                if self.selected_image_index in range(len(image_files)-1):
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
        except Exception as e:
            self.outofIndex(e)        

    def output_next_image_clicked_Segment_model(self):
        """ """
        # get first image file and show it
        try:
            self.output_image_path_ = self.lineEdit_2.text()
            output_image_path_ = glob.glob(os.path.join(str(self.output_image_path_), "Predict_output/Output_csv", "*"))

            image_files_2 = []
            for folder in output_image_path_:
                for f in glob.glob(folder+'/*.jpg'):
                    image_files_2.append(f)

            if bool(output_image_path_) is True:

                if self.selected_image_index in range(len(image_files_2)-1):
                    # If reached the last image, set index to start over
                    if self.selected_image_index == len(image_files_2) - 1:
                        self.selected_image_index = 0
                    else:
                        self.selected_image_index += 1
                else:
                    self.selected_image_index -= 1
                scene = QtWidgets.QGraphicsScene(StoManager1)
                pixmap = QPixmap(image_files_2[self.selected_image_index])
                item = QtWidgets.QGraphicsPixmapItem(pixmap)
                scene.addItem(item)
                self.graphicsView_2.setScene(scene)
                #self.selected_image_index += 1

            else:
                self.messagebox_define_output_path()
                pass
        except Exception as e:
            self.outofIndex(e)        

# previous image
    def input_previous_image_clicked(self):
        """ """
        try:
            self.input_image_path_ = self.lineEdit.text()
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:

                if self.selected_image_index == 0:
                    self.selected_image_index = len(image_files_2) - 1
                else:
                    self.selected_image_index -= 1
                scene = QtWidgets.QGraphicsScene(StoManager1)
                pixmap = QPixmap(image_files_2[self.selected_image_index])
                item = QtWidgets.QGraphicsPixmapItem(pixmap)
                scene.addItem(item)
                self.graphicsView.setScene(scene)
            else:
                self.messagebox_define_input_path()
                pass
        except Exception as e:
            self.outofIndex(e)

# previous image
    def output_previous_image_clicked_Box_model(self):
        """ """
        try:
            self.output_image_path_ = self.lineEdit_2.text()
            image_files = glob.glob(self.output_image_path_ + "/" + '*jpg')
            input_image_path = self.input_image_path_
            # Specify the file extensions you want to include
            allowed_extensions = ['jpg', 'png', 'tif','jpeg']
            # Create a list of image files with the specified extensions
            image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
            image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
            has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

            if has_image_files:

                if self.selected_image_index == 0:
                    self.selected_image_index = 0
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
        except Exception as e:
            self.outofIndex(e)

    def output_previous_image_clicked_Segment_model(self):
        """ """
        try:
            self.output_image_path_ = self.lineEdit_2.text()
            output_image_path_ = glob.glob(os.path.join(str(self.output_image_path_), "Predict_output/Output_csv", "*"))

            image_files_2 = []
            for folder in output_image_path_:
                for f in glob.glob(folder+'/*.jpg'):
                    image_files_2.append(f)

            if bool(output_image_path_) is True:

                if self.selected_image_index == 0:
                    self.selected_image_index = len(image_files_2)-1
                else:
                    self.selected_image_index -= 1
                scene = QtWidgets.QGraphicsScene(StoManager1)
                pixmap = QPixmap(image_files_2[self.selected_image_index])
                item = QtWidgets.QGraphicsPixmapItem(pixmap)
                scene.addItem(item)
                self.graphicsView_2.setScene(scene)
            else:
                self.messagebox_define_output_path()
                pass
        except Exception as e:
            self.outofIndex(e)

###### Main components ######

    # define a function to normalize the file names
    def NormalizeFileNames(self):
        """ This function is now only for my specific project,
         it automatically allocate and group data based on Site, Block, and Clone """
        self.input_image_path_ = self.lineEdit.text()
        Output_path = self.lineEdit_2.text()
        input_image_path = self.input_image_path_
        # Specify the file extensions you want to include
        allowed_extensions = ['jpg', 'png', 'tif','jpeg']
        # Create a list of image files with the specified extensions
        image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
        image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
        has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

        if has_image_files:        
            file_path = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))][2]
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

                for file_path in [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]:
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

    def Box_model(self):  ## the main funtion with the yolov3 model we trained to detect our stomata
        """ """
        self.confidence = self.lineEdit_5.text()
        self.input_image_path_ = self.lineEdit.text()
        Output_path = self.lineEdit_2.text()
        input_image_path = self.input_image_path_
        
        self.p = self.lineEdit_4.text()
        # Specify the file extensions you want to include
        allowed_extensions = ['jpg', 'png', 'tif','jpeg']
        # Create a list of image files with the specified extensions
        image_files_2 = [file for ext in allowed_extensions for file in glob.glob(os.path.join(input_image_path, f'*.{ext}'))]
        image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
        has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

        if has_image_files: 

            self.img_num = 1

            if self.p =="Input pixels in 0.1 mm, the default is: 465" or self.p =="" or self.p ==" ":
                self.p = 465
            else: 
                self.p = int(self.lineEdit_4.text())

            pixel = float(self.p)

            # Load Yolo
            net = cv2.dnn.readNet("yolov3_training_last.weights", "yolov3_testing.cfg")

            # Name custom object
            classes = ["whole_stomata", "stomata"]

            # Images path
            images_path = image_files_2

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
                self.confidence = self.lineEdit_5.text()
            
                if self.confidence =="Confidence threshold for detection the default is 0.25" or self.confidence =="" or self.confidence ==" ":
                    self.confidence = 0.25
                else: 
                    self.confidence = float(self.lineEdit_5.text())
                
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
                        if confidence > self.confidence:
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
                                ((w * h) * 0.6878 + 806)*(10000/(pixel*pixel)))  ## Build linear regression model for adjusted area of whole_stomata
                            list_of_whole_stomatal_area.append(((w * h) * 0.6878 + 806)*(10000/(pixel*pixel)))
                        elif label == "stomata":  ## Build linear regression model for adjusted area of stomata
                            number_of_stomata.append(class_ids[i])
                            list_of_all_stomata_areas.append(
                                ((w * h + 116.08) / 1.7684)*(10000/(pixel*pixel)))
                            
                        orientation = math.log(w/h)*(-92.2325)+44.5222
                        
                        if orientation>=0:
                            orientations.append(orientation)
                        else:
                            orientations.append(orientation+180)

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
                try:
                    list_of_whole_stomatal_area_ratio = sum((area/(10000/(pixel*pixel))) for area in list_of_whole_stomatal_area) / (
                            np.mean(list_of_image_width) * np.mean(list_of_image_height))
                    heights = list_of_height
                    widths = list_of_width
                    image_name = img_path
                    image_width = list_of_image_width
                    image_height = list_of_image_height
                    all_stomata_areas = [(float(k)) for k in list_of_all_stomata_areas]
                    Whole_stomata_density=number_of_whole_stomata
                    image_name = {"Labels": labels, "Width_(pixels)": widths, "Height_(pixels)": heights,
                                "Orientation": orientations,
                                "Num_of_Stomata": len(number_of_stomata), "Num_of_whole_stomata": len(number_of_whole_stomata),
                                "Width_of_image_(pixels)": image_width, "Height_of_image_(pixels)": image_height,
                                "All_sotmata_area_(mum2)": all_stomata_areas,
                                "Whole_stomata_area_ratio": list_of_whole_stomatal_area_ratio,
                                "Whole_stomata_density": len(Whole_stomata_density)/(height*width/(pixel*10.0)**2)}
                except RuntimeWarning():
                    self.runtimeWarning()
                df = pd.DataFrame(image_name)

                # print(image_name)
                # df.to_csv('width_height.csv')
                img_path = str(img_path)
                new_img_path = img_path.replace("\\", "")
                new_img_path = img_path.replace(".jpg", "")
                print(new_img_path)
                
                # check if the files are currently opening ? if yes, create a copy, if not replace the older ones
                try: 
                    file = df.to_csv(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".csv", sep=',', index=True)
                    cv2.imwrite(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".jpg", img)
                except OSError:
                    self.closeExcel()

            self.show_info_messagebox()
        else:
            self.messagebox_define_input_path_2()
            pass
    # Define a function to analyze the  stomatal data

    def Stomata_group_analysis(self):
        """ This function is designed for my specific project which has Site, Block, and Clone information"""
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
                if single_csv_file.shape[0]>=4:                
                    file_name = self.file_path[len(self.output_image_path_) + 1:-4]  # Extract the file name from the file path
                    Split_name = file_name.split(",")                 
                    if len(Split_name)>=5:
                                        
                        for self.file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                            single_csv_file = pd.read_csv(self.file_path, low_memory=False)  # Read the csv and assign it to a variable
                            file_name = self.file_path[len(self.output_image_path_) + 1:-4]  # Extract the file name from the file path
                            Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
                            
                            try:
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
                                Whole_stomata_areas_max = max(Whole_stomata_areas)
                                Whole_stomata_areas_min = min(Whole_stomata_areas)
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
                                WST_Density.append(int(np.mean(Whole_stomata_density)))
                                WST_Number.append(np.mean(Whole_stomata_number))
                                WST_Area.append(np.mean(Whole_stomata_areas))
                                WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))
                                self.progressBar.setValue(int((self.img_num / len(glob.glob(self.output_image_path_ + '/' + '*.csv'))*100)))
                                QtWidgets.QApplication.processEvents()

                                self.img_num +=1
                            except KeyError:
                                self.keyError()
                                self.img_num +=2                            
                            
                        else:
                            self.img_num +=1
                            pass                        
                            
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

    def Stomata_group_analysis_seg(self):
        """ This function is designed for my specific project which has Site, Block, and Clone information"""
        # Create empty lists to hold the values that we are going to extract
        self.output_folder = self.lineEdit_2.text()

        output_image_path_ = os.path.join(str(self.output_folder), "Predict_output/Output_csv")

        if bool(glob.glob(output_image_path_ + "/" + '*csv')) is True:

            No_wst_mean = []
            No_wst_median = []
            No_wst_min = []
            No_wst_max = []

            box_w_wst_mean = []
            box_w_wst_median = []
            box_w_wst_min = []
            box_w_wst_max = []

            box_h_wst_mean = []
            box_h_wst_median = []
            box_h_wst_min = []
            box_h_wst_max = []

            area_wst_mean = []
            area_wst_median = []
            area_wst_min = []
            area_wst_max = []

            width_wst_mean = []
            width_wst_median = []
            width_wst_min = []
            width_wst_max = []

            length_wst_mean = []
            length_wst_median = []
            length_wst_min = []
            length_wst_max = []

            var_area_wst_mean = []
            var_area_wst_median = []
            var_area_wst_min = []
            var_area_wst_max = []

            var_width_wst_mean = []
            var_width_wst_median = []
            var_width_wst_min = []
            var_width_wst_max = []

            var_length_wst_mean = []
            var_length_wst_median = []
            var_length_wst_min = []
            var_length_wst_max = []    

            ## for stomata
            No_st_mean = []
            No_st_median = []
            No_st_min = []
            No_st_max = []

            box_w_st_mean = []
            box_w_st_median = []
            box_w_st_min = []
            box_w_st_max = []

            box_h_st_mean = []
            box_h_st_median = []
            box_h_st_min = []
            box_h_st_max = []

            area_st_mean = []
            area_st_median = []
            area_st_min = []
            area_st_max = []

            width_st_mean = []
            width_st_median = []
            width_st_min = []
            width_st_max = []

            length_st_mean = []
            length_st_median = []
            length_st_min = []
            length_st_max = []

            var_area_st_mean = []
            var_area_st_median = []
            var_area_st_min = []
            var_area_st_max = []

            var_width_st_mean = []
            var_width_st_median = []
            var_width_st_min = []
            var_width_st_max = []

            var_length_st_mean = []
            var_length_st_median = []
            var_length_st_min = []
            var_length_st_max = []   

            ## for guard cell
            guardCell_length_mean = []  
            guardCell_length_median = []    
            guardCell_length_min = [] 
            guardCell_length_max = []    

            guardCell_width_mean = []
            guardCell_width_median = []
            guardCell_width_min = []
            guardCell_width_max = []

            guardCell_area_mean = []
            guardCell_area_median = []
            guardCell_area_min = []
            guardCell_area_max = []

            guardCell_angle_mean = []
            guardCell_angle_median = []
            guardCell_angle_min = []
            guardCell_angle_max = []

            var_width_guardCell_mean = []
            var_width_guardCell_median = []
            var_width_guardCell_min = []
            var_width_guardCell_max = []

            var_length_guardCell_mean = []
            var_length_guardCell_median = []
            var_length_guardCell_min = []
            var_length_guardCell_max = []

            wst_density_mean = []
            wst_density_median = []
            wst_density_min = []
            wst_density_max = []

            ratio_area_st_gc_mean = []
            ratio_area_st_gc_median = []
            ratio_area_st_gc_min = []
            ratio_area_st_gc_max = []

            ratio_area_to_img_mean = []
            ratio_area_to_img_median = []
            ratio_area_to_img_min = []
            ratio_area_to_img_max = []

            var_angle_mean = []
            var_angle_median = []
            var_angle_min = []
            var_angle_max = []

            SEve_mean = []
            SEve_median = []
            SEve_min = []
            SEve_max = []            

            SDiv_mean = []
            SDiv_median = []
            SDiv_min = []
            SDiv_max = []     

            SAgg_mean = []
            SAgg_median = []
            SAgg_min = []
            SAgg_max = []            

            Filename = []            

            Site = []
            Block = []
            Clone = []
            Month = []
            Year = []

            # Using for loop to go through all csv files
            if bool(glob.glob(output_image_path_ + '/' + 'Statistics.csv')) is True:
                os.remove(glob.glob(output_image_path_ + '/' + 'Statistics.csv')[0])
            else:
                self.img_num =1
                # for file_path in glob.glob(self.output_image_path_ + '/' + '*.csv'):
                file_path = glob.glob(output_image_path_ + '/' + '*.csv')[1]
                single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable
                if single_csv_file.shape[0]>=4:                
                    file_name = file_path[len(output_image_path_) + 1:-4]  # Extract the file name from the file path
                    Split_name = file_name.split(",")                 
                    if len(Split_name)>=5:
                                        
                        for file_path in glob.glob(output_image_path_ + '/' + '*.csv'):
                            single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable
                            file_name = file_path[len(output_image_path_) + 1:-4]  # Extract the file name from the file path
                            Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
                            
                            try:

                                Filename.append(",".join([str(item) for item in Split_name]))
                                
                                Site.append(Split_name[0])  # Add the site to the Site list
                                Block.append(Split_name[1])
                                Clone.append(Split_name[2])
                                Month.append(Split_name[3])
                                Year.append(Split_name[4])

                                # Extract the parameters from each single csv files
                                number_wst = single_csv_file["number_wst"]
                                box_w_wst = single_csv_file["box_w_wst"]
                                box_h_wst = single_csv_file["box_h_wst"]  
                                area_wst = single_csv_file["area_wst"]
                                width_wst = single_csv_file["width_wst"]
                                length_wst = single_csv_file["length_wst"]
                                var_area_wst = single_csv_file["var_area_wst"]
                                var_width_wst = single_csv_file["var_width_wst"]
                                var_length_wst = single_csv_file["var_length_wst"]
                                
                                number_st = single_csv_file["number_st"]
                                box_w_st = single_csv_file["box_w_st"]
                                box_h_st = single_csv_file["box_h_st"]  
                                area_st = single_csv_file["area_st"]
                                width_st = single_csv_file["width_st"]
                                length_st = single_csv_file["length_st"]
                                var_area_st = single_csv_file["var_area_st"]
                                var_width_st = single_csv_file["var_width_st"]
                                var_length_st = single_csv_file["var_length_st"]

                                guardCell_length = single_csv_file["guardCell_length"]
                                guardCell_width = single_csv_file["guardCell_width"]
                                guardCell_area = single_csv_file["guardCell_area"]
                                guardCell_angle = single_csv_file["guardCell_angle"]
                                var_angle = single_csv_file["var_angle"]
                                var_width_guardCell = single_csv_file["var_width_guardCell"]
                                var_length_guardCell = single_csv_file["var_length_guardCell"]
                                wst_density = single_csv_file["wst_density"]
                                ratio_area_st_gc = single_csv_file["ratio_area_st_to_gc"]
                                ratio_area_to_img = single_csv_file["ratio_area_to_img"]

                                SEve = single_csv_file["SEve"]
                                SDiv = single_csv_file["SDiv"]
                                SAgg = single_csv_file["SAgg"]                                

                                # print(type(box_w_wst))

                                # Remove the outliers before calculating box_w_wst variance
                                Q1_box_w_wst = np.percentile(box_w_wst, 5, method = 'midpoint')
                                Q3_box_w_wst = np.percentile(box_w_wst, 95, method = 'midpoint')
                                IQR_box_w_wst = Q3_box_w_wst - Q1_box_w_wst
                                upper = np.where(box_w_wst >= (Q3_box_w_wst+1.5*IQR_box_w_wst))
                                lower = np.where(box_w_wst <= (Q1_box_w_wst-1.5*IQR_box_w_wst))
                                box_w_wst.drop(upper[0], inplace = True)
                                box_w_wst.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating box_h_wst variance
                                Q1_box_h_wst = np.percentile(box_h_wst, 2.5, method = 'midpoint')
                                Q3_box_h_wst = np.percentile(box_h_wst, 97.5,method = 'midpoint')
                                IQR_box_h_wst = Q3_box_h_wst - Q1_box_h_wst
                                upper = np.where(box_h_wst >= (Q3_box_h_wst+1.5*IQR_box_h_wst))
                                lower = np.where(box_h_wst <= (Q1_box_h_wst-1.5*IQR_box_h_wst))
                                box_h_wst.drop(upper[0], inplace = True)
                                box_h_wst.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating area_wst variance
                                Q1_area_wst = np.percentile(area_wst, 2.5, method = 'midpoint')
                                Q3_area_wst = np.percentile(area_wst, 97.5,method = 'midpoint')
                                IQR_area_wst = Q3_area_wst - Q1_area_wst
                                upper = np.where(area_wst >= (Q3_area_wst+1.5*IQR_area_wst))
                                lower = np.where(area_wst <= (Q1_area_wst-1.5*IQR_area_wst))
                                area_wst.drop(upper[0], inplace = True)
                                area_wst.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating width_wst variance
                                Q1_width_wst = np.percentile(width_wst, 2.5, method = 'midpoint')
                                Q3_width_wst = np.percentile(width_wst, 97.5,method = 'midpoint')
                                IQR_width_wst = Q3_width_wst - Q1_width_wst
                                upper = np.where(width_wst >= (Q3_width_wst+1.5*IQR_width_wst))
                                lower = np.where(width_wst <= (Q1_width_wst-1.5*IQR_width_wst))
                                width_wst.drop(upper[0], inplace = True)
                                width_wst.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating length_wst variance
                                Q1_length_wst = np.percentile(length_wst, 2.5, method = 'midpoint')
                                Q3_length_wst = np.percentile(length_wst, 97.5,method = 'midpoint')
                                IQR_length_wst = Q3_length_wst - Q1_length_wst
                                upper = np.where(length_wst >= (Q3_length_wst+1.5*IQR_length_wst))
                                lower = np.where(length_wst <= (Q1_length_wst-1.5*IQR_length_wst))
                                length_wst.drop(upper[0], inplace = True)
                                length_wst.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating var_area_wst variance
                                # Q1_var_area_wst = np.percentile(var_area_wst, 5, method = 'midpoint')
                                # Q3_var_area_wst = np.percentile(var_area_wst, 95,method = 'midpoint')
                                # IQR_var_area_wst = Q3_var_area_wst - Q1_var_area_wst
                                # upper = np.where(var_area_wst >= (Q3_var_area_wst+1.5*IQR_var_area_wst))
                                # lower = np.where(var_area_wst <= (Q1_var_area_wst-1.5*IQR_var_area_wst))
                                # var_area_wst.drop(upper[0], inplace = True)
                                # var_area_wst.drop(lower[0], inplace = True)
                                var_area_wst = var_area_wst[-1:]

                                # Remove the outliers before calculating var_width_wst variance
                                # Q1_var_width_wst = np.percentile(var_width_wst, 5, method = 'midpoint')
                                # Q3_var_width_wst = np.percentile(var_width_wst, 95,method = 'midpoint')
                                # IQR_var_width_wst = Q3_var_width_wst - Q1_var_width_wst
                                # upper = np.where(var_width_wst >= (Q3_var_width_wst+1.5*IQR_var_width_wst))
                                # lower = np.where(var_width_wst <= (Q1_var_width_wst-1.5*IQR_var_width_wst))
                                # var_width_wst.drop(upper[0], inplace = True)
                                # var_width_wst.drop(lower[0], inplace = True) 
                                var_width_wst = var_width_wst[-1:]                

                                var_angle = var_angle[-1:]   
                                # Remove the outliers before calculating var_length_wst variance
                                # Q1_var_length_wst = np.percentile(var_length_wst, 5, method = 'midpoint')
                                # Q3_var_length_wst = np.percentile(var_length_wst, 95,method = 'midpoint')
                                # IQR_var_length_wst = Q3_var_length_wst - Q1_var_length_wst
                                # upper = np.where(var_length_wst >= (Q3_var_length_wst+1.5*IQR_var_length_wst))
                                # lower = np.where(var_length_wst <= (Q1_var_length_wst-1.5*IQR_var_length_wst))
                                # var_length_wst.drop(upper[0], inplace = True)
                                # var_length_wst.drop(lower[0], inplace = True)  
                                var_length_wst = var_length_wst[-1:] 

                                # Remove the outliers before calculating box_w_st variance
                                Q1_box_w_st = np.percentile(box_w_st, 2.5, method = 'midpoint')
                                Q3_box_w_st = np.percentile(box_w_st, 97.5,method = 'midpoint')
                                IQR_box_w_st = Q3_box_w_st - Q1_box_w_st
                                upper = np.where(box_w_st >= (Q3_box_w_st+1.5*IQR_box_w_st))
                                lower = np.where(box_w_st <= (Q1_box_w_st-1.5*IQR_box_w_st))
                                box_w_st.drop(upper[0], inplace = True)
                                box_w_st.drop(lower[0], inplace = True)                               

                                # Remove the outliers before calculating box_h_st variance
                                Q1_box_h_st = np.percentile(box_h_st, 2.5, method = 'midpoint')
                                Q3_box_h_st = np.percentile(box_h_st, 97.5,method = 'midpoint')
                                IQR_box_h_st = Q3_box_h_st - Q1_box_h_st
                                upper = np.where(box_h_st >= (Q3_box_h_st+1.5*IQR_box_h_st))
                                lower = np.where(box_h_st <= (Q1_box_h_st-1.5*IQR_box_h_st))
                                box_h_st.drop(upper[0], inplace = True)
                                box_h_st.drop(lower[0], inplace = True)   

                                # Remove the outliers before calculating area_st variance
                                Q1_area_st = np.percentile(area_st, 2.5, method = 'midpoint')
                                Q3_area_st = np.percentile(area_st, 97.5,method = 'midpoint')
                                IQR_area_st = Q3_area_st - Q1_area_st
                                upper = np.where(area_st >= (Q3_area_st+1.5*IQR_area_st))
                                lower = np.where(area_st <= (Q1_area_st-1.5*IQR_area_st))
                                area_st.drop(upper[0], inplace = True)
                                area_st.drop(lower[0], inplace = True)   

                                # Remove the outliers before calculating width_st variance
                                Q1_width_st = np.percentile(width_st, 2.5, method = 'midpoint')
                                Q3_width_st = np.percentile(width_st, 97.5,method = 'midpoint')
                                IQR_width_st = Q3_width_st - Q1_width_st
                                upper = np.where(width_st >= (Q3_width_st+1.5*IQR_width_st))
                                lower = np.where(width_st <= (Q1_width_st-1.5*IQR_width_st))
                                width_st.drop(upper[0], inplace = True)
                                width_st.drop(lower[0], inplace = True)  

                                # Remove the outliers before calculating length_st variance
                                Q1_length_st = np.percentile(length_st, 2.5, method = 'midpoint')
                                Q3_length_st = np.percentile(length_st, 97.5,method = 'midpoint')
                                IQR_length_st = Q3_length_st - Q1_length_st
                                upper = np.where(length_st >= (Q3_length_st+1.5*IQR_length_st))
                                lower = np.where(length_st <= (Q1_length_st-1.5*IQR_length_st))
                                length_st.drop(upper[0], inplace = True)
                                length_st.drop(lower[0], inplace = True)

                                # Remove the outliers before calculating var_area_st variance
                                # Q1_var_area_st = np.percentile(var_area_st, 5, method = 'midpoint')
                                # Q3_var_area_st = np.percentile(var_area_st, 95,method = 'midpoint')
                                # IQR_var_area_st = Q3_var_area_st - Q1_var_area_st
                                # upper = np.where(var_area_st >= (Q3_var_area_st+1.5*IQR_var_area_st))
                                # lower = np.where(var_area_st <= (Q1_var_area_st-1.5*IQR_var_area_st))
                                # var_area_st.drop(upper[0], inplace = True)
                                # var_area_st.drop(lower[0], inplace = True)
                                var_area_st = var_area_st[-1:] 

                                # Remove the outliers before calculating var_width_st variance
                                # Q1_var_width_st = np.percentile(var_width_st, 5, method = 'midpoint')
                                # Q3_var_width_st = np.percentile(var_width_st, 95,method = 'midpoint')
                                # IQR_var_width_st = Q3_var_width_st - Q1_var_width_st
                                # upper = np.where(var_width_st >= (Q3_var_width_st+1.5*IQR_var_width_st))
                                # lower = np.where(var_width_st <= (Q1_var_width_st-1.5*IQR_var_width_st))
                                # var_width_st.drop(upper[0], inplace = True)
                                # var_width_st.drop(lower[0], inplace = True) 
                                var_width_st = var_width_st[-1:]                 

                                # Remove the outliers before calculating var_length_st variance
                                # Q1_var_length_st = np.percentile(var_length_st, 5, method = 'midpoint')
                                # Q3_var_length_st = np.percentile(var_length_st, 95,method = 'midpoint')
                                # IQR_var_length_st = Q3_var_length_st - Q1_var_length_st
                                # upper = np.where(var_length_st >= (Q3_var_length_st+1.5*IQR_var_length_st))
                                # lower = np.where(var_length_st <= (Q1_var_length_st-1.5*IQR_var_length_st))
                                # var_length_st.drop(upper[0], inplace = True)
                                # var_length_st.drop(lower[0], inplace = True) 
                                var_length_st = var_length_st[-1:]                

                                # Remove the outliers before calculating guardCell_length variance
                                Q1_guardCell_length = np.percentile(guardCell_length, 2.5, method = 'midpoint')
                                Q3_guardCell_length = np.percentile(guardCell_length, 97.5,method = 'midpoint')
                                IQR_guardCell_length = Q3_guardCell_length - Q1_guardCell_length
                                upper = np.where(guardCell_length >= (Q3_guardCell_length+1.5*IQR_guardCell_length))
                                lower = np.where(guardCell_length <= (Q1_guardCell_length-1.5*IQR_guardCell_length))
                                guardCell_length.drop(upper[0], inplace = True)
                                guardCell_length.drop(lower[0], inplace = True) 
                                
                                # Remove the outliers before calculating guardCell_width variance
                                Q1_guardCell_width = np.percentile(guardCell_width, 2.5, method = 'midpoint')
                                Q3_guardCell_width = np.percentile(guardCell_width, 97.5,method = 'midpoint')
                                IQR_guardCell_width = Q3_guardCell_width - Q1_guardCell_width
                                upper = np.where(guardCell_width >= (Q3_guardCell_width+1.5*IQR_guardCell_width))
                                lower = np.where(guardCell_width <= (Q1_guardCell_width-1.5*IQR_guardCell_width))
                                guardCell_width.drop(upper[0], inplace = True)
                                guardCell_width.drop(lower[0], inplace = True) 

                                # Remove the outliers before calculating guardCell_area variance
                                Q1_guardCell_area = np.percentile(guardCell_area, 2.5, method = 'midpoint')
                                Q3_guardCell_area = np.percentile(guardCell_area, 97.5,method = 'midpoint')
                                IQR_guardCell_area = Q3_guardCell_area - Q1_guardCell_area
                                upper = np.where(guardCell_area >= (Q3_guardCell_area+1.5*IQR_guardCell_area))
                                lower = np.where(guardCell_area <= (Q1_guardCell_area-1.5*IQR_guardCell_area))
                                guardCell_area.drop(upper[0], inplace = True)
                                guardCell_area.drop(lower[0], inplace = True) 

                                # Remove the outliers before calculating var_angle variance
                                # Q1_var_angle = np.percentile(var_angle, 5, method = 'midpoint')
                                # Q3_var_angle = np.percentile(var_angle, 95,method = 'midpoint')
                                # IQR_var_angle = Q3_var_angle - Q1_var_angle
                                # upper = np.where(var_angle >= (Q3_var_angle+1.5*IQR_var_angle))
                                # lower = np.where(var_angle <= (Q1_var_angle-1.5*IQR_var_angle))
                                # var_angle.drop(upper[0], inplace = True)
                                # var_angle.drop(lower[0], inplace = True) 

                                # Remove the outliers before calculating var_width_guardCell variance
                                # Q1_var_width_guardCell = np.percentile(var_width_guardCell, 5, method = 'midpoint')
                                # Q3_var_width_guardCell = np.percentile(var_width_guardCell, 95,method = 'midpoint')
                                # IQR_var_width_guardCell = Q3_var_width_guardCell - Q1_var_width_guardCell
                                # upper = np.where(var_width_guardCell >= (Q3_var_width_guardCell+1.5*IQR_var_width_guardCell))
                                # lower = np.where(var_width_guardCell <= (Q1_var_width_guardCell-1.5*IQR_var_width_guardCell))
                                # var_width_guardCell.drop(upper[0], inplace = True)
                                # var_width_guardCell.drop(lower[0], inplace = True)
                                var_width_guardCell = var_width_guardCell[-1:]                  

                                # Remove the outliers before calculating var_length_guardCell variance
                                # Q1_var_length_guardCell = np.percentile(var_length_guardCell, 5, method = 'midpoint')
                                # Q3_var_length_guardCell = np.percentile(var_length_guardCell, 95,method = 'midpoint')
                                # IQR_var_length_guardCell = Q3_var_length_guardCell - Q1_var_length_guardCell
                                # upper = np.where(var_length_guardCell >= (Q3_var_length_guardCell+1.5*IQR_var_length_guardCell))
                                # lower = np.where(var_length_guardCell <= (Q1_var_length_guardCell-1.5*IQR_var_length_guardCell))
                                # var_length_guardCell.drop(upper[0], inplace = True)
                                # var_length_guardCell.drop(lower[0], inplace = True) 
                                var_length_guardCell = var_length_guardCell[-1:] 

                                # Remove the outliers before calculating ratio_area_st_gc variance
                                Q1_ratio_area_st_gc = np.percentile(ratio_area_st_gc, 2.5, method = 'midpoint')
                                Q3_ratio_area_st_gc = np.percentile(ratio_area_st_gc, 97.5,method = 'midpoint')
                                IQR_ratio_area_st_gc = Q3_ratio_area_st_gc - Q1_ratio_area_st_gc
                                upper = np.where(ratio_area_st_gc >= (Q3_ratio_area_st_gc+1.5*IQR_ratio_area_st_gc))
                                lower = np.where(ratio_area_st_gc <= (Q1_ratio_area_st_gc-1.5*IQR_ratio_area_st_gc))
                                ratio_area_st_gc.drop(upper[0], inplace = True)
                                ratio_area_st_gc.drop(lower[0], inplace = True) 

                                # Remove the outliers before calculating ratio_area_to_img variance
                                ratio_area_to_img = ratio_area_to_img[-1:] 

                                # calculate the median, mean, variance of each leaf stomata number, stomata area

                                No_wst_mean_ = np.mean(number_wst)
                                No_wst_median_ = np.median(number_wst)
                                No_wst_min_ = min(number_wst)
                                No_wst_max_ = max(number_wst)

                                box_w_wst_mean_ = np.mean(box_w_wst)
                                box_w_wst_median_ = np.median(box_w_wst)
                                box_w_wst_min_ = min(box_w_wst)
                                box_w_wst_max_ = max(box_w_wst)

                                box_h_wst_mean_ = np.mean(box_h_wst)
                                box_h_wst_median_ = np.median(box_h_wst)
                                box_h_wst_min_ = min(box_h_wst)
                                box_h_wst_max_ = max(box_h_wst)

                                area_wst_mean_ = np.mean(area_wst)
                                area_wst_median_ = np.median(area_wst)
                                area_wst_min_ = min(area_wst)
                                area_wst_max_ = max(area_wst)

                                width_wst_mean_ = np.mean(width_wst)
                                width_wst_median_ = np.median(width_wst)
                                width_wst_min_ = min(width_wst)
                                width_wst_max_ = max(width_wst)

                                length_wst_mean_ = np.mean(length_wst)
                                length_wst_median_ = np.median(length_wst)
                                length_wst_min_ = min(length_wst)
                                length_wst_max_ = max(length_wst)

                                var_area_wst_mean_ = np.mean(var_area_wst)
                                var_area_wst_median_ = np.median(var_area_wst)
                                var_area_wst_min_ = min(var_area_wst)
                                var_area_wst_max_ = max(var_area_wst)

                                var_width_wst_mean_ = np.mean(var_width_wst)
                                var_width_wst_median_ = np.median(var_width_wst)
                                var_width_wst_min_ = min(var_width_wst)
                                var_width_wst_max_ = max(var_width_wst)

                                var_length_wst_mean_ = np.mean(var_length_wst)
                                var_length_wst_median_ = np.median(var_length_wst)
                                var_length_wst_min_ = min(var_length_wst)
                                var_length_wst_max_ = max(var_length_wst)  

                                ## for stomata
                                No_st_mean_ = np.mean(number_st)
                                No_st_median_ = np.median(number_st)
                                No_st_min_ = min(number_st)
                                No_st_max_ = max(number_st)

                                box_w_st_mean_ = np.mean(box_w_st)
                                box_w_st_median_ = np.median(box_w_st)
                                box_w_st_min_ = min(box_w_st)
                                box_w_st_max_ = max(box_w_st)

                                box_h_st_mean_ = np.mean(box_h_st)
                                box_h_st_median_ = np.median(box_h_st)
                                box_h_st_min_ = min(box_h_st)
                                box_h_st_max_ = max(box_h_st)

                                area_st_mean_ = np.mean(area_st)
                                area_st_median_ = np.median(area_st)
                                area_st_min_ = min(area_st)
                                area_st_max_ = max(area_st)

                                width_st_mean_ = np.mean(width_st)
                                width_st_median_ = np.median(width_st)
                                width_st_min_ = min(width_st)
                                width_st_max_ = max(width_st)

                                length_st_mean_ = np.mean(length_st)
                                length_st_median_ = np.median(length_st)
                                length_st_min_ = min(length_st)
                                length_st_max_ = max(length_st)

                                var_area_st_mean_ = np.mean(var_area_st)
                                var_area_st_median_ = np.median(var_area_st)
                                var_area_st_min_ = min(var_area_st)
                                var_area_st_max_ = max(var_area_st)

                                var_width_st_mean_ = np.mean(var_width_st)
                                var_width_st_median_ = np.median(var_width_st)
                                var_width_st_min_ = min(var_width_st)
                                var_width_st_max_ = max(var_width_st)

                                var_length_st_mean_ = np.mean(var_length_st)
                                var_length_st_median_ = np.median(var_length_st)
                                var_length_st_min_ = min(var_length_st)
                                var_length_st_max_ = max(var_length_st)  

                                ## for guard cell
                                guardCell_length_mean_ = np.mean(guardCell_length)  
                                guardCell_length_median_ = np.median(guardCell_length)  
                                guardCell_length_min_ = min(guardCell_length)
                                guardCell_length_max_ = max(guardCell_length)    

                                guardCell_width_mean_ = np.mean(guardCell_width)
                                guardCell_width_median_ = np.median(guardCell_width)
                                guardCell_width_min_ = min(guardCell_width)
                                guardCell_width_max_ = max(guardCell_width)

                                guardCell_area_mean_ = np.mean(guardCell_area)
                                guardCell_area_median_ = np.median(guardCell_area)
                                guardCell_area_min_ = min(guardCell_area)
                                guardCell_area_max_ = max(guardCell_area)

                                guardCell_angle_mean_ = np.mean(guardCell_angle)
                                guardCell_angle_median_ = np.median(guardCell_angle)
                                guardCell_angle_min_ = min(guardCell_angle)
                                guardCell_angle_max_ = max(guardCell_angle)

                                var_width_guardCell_mean_ = np.mean(var_width_guardCell)
                                var_width_guardCell_median_ = np.median(var_width_guardCell)
                                var_width_guardCell_min_ = min(var_width_guardCell)
                                var_width_guardCell_max_ = max(var_width_guardCell)

                                var_length_guardCell_mean_ = np.mean(var_length_guardCell)
                                var_length_guardCell_median_ = np.median(var_length_guardCell)
                                var_length_guardCell_min_ = min(var_length_guardCell)
                                var_length_guardCell_max_ = max(var_length_guardCell)

                                wst_density_mean_ = np.mean(wst_density)
                                wst_density_median_ = np.median(wst_density)
                                wst_density_min_ = min(wst_density)
                                wst_density_max_ = max(wst_density)

                                ratio_area_st_gc_mean_ = np.mean(ratio_area_st_gc)
                                ratio_area_st_gc_median_ = np.median(ratio_area_st_gc)
                                ratio_area_st_gc_min_ = min(ratio_area_st_gc)
                                ratio_area_st_gc_max_ = max(ratio_area_st_gc)

                                ratio_area_to_img_mean_ = np.mean(ratio_area_to_img)
                                ratio_area_to_img_median_ = np.median(ratio_area_to_img)
                                ratio_area_to_img_min_ = min(ratio_area_to_img)
                                ratio_area_to_img_max_ = max(ratio_area_to_img)

                                var_angle_mean_ = np.mean(var_angle)
                                var_angle_median_ = np.median(var_angle)
                                var_angle_min_ = min(var_angle)
                                var_angle_max_ = max(var_angle)

                                SEve_mean_ = np.mean(SEve)
                                SEve_median_ = np.median(SEve)
                                SEve_min_ = min(SEve)
                                SEve_max_ = max(SEve)

                                SDiv_mean_ = np.mean(SDiv)
                                SDiv_median_ = np.median(SDiv)
                                SDiv_min_ = min(SDiv)
                                SDiv_max_ = max(SDiv)

                                SAgg_mean_ = np.mean(SAgg)
                                SAgg_median_ = np.median(SAgg)
                                SAgg_min_ = min(SAgg)
                                SAgg_max_ = max(SAgg)                                

                                ## append these measurements to the lists
                                # for whole_stomata
                                No_wst_mean.append(No_wst_mean_)
                                No_wst_median.append(No_wst_median_)
                                No_wst_min.append(No_wst_min_)
                                No_wst_max.append(No_wst_max_)

                                box_w_wst_mean.append(box_w_wst_mean_)
                                box_w_wst_median.append(box_w_wst_median_)
                                box_w_wst_min.append(box_w_wst_min_)
                                box_w_wst_max.append(box_w_wst_max_)

                                box_h_wst_mean.append(box_h_wst_mean_)
                                box_h_wst_median.append(box_h_wst_median_)
                                box_h_wst_min.append(box_h_wst_min_)
                                box_h_wst_max.append(box_h_wst_max_)

                                area_wst_mean.append(area_wst_mean_)
                                area_wst_median.append(area_wst_median_)
                                area_wst_min.append(area_wst_min_)
                                area_wst_max.append(area_wst_max_)

                                width_wst_mean.append(width_wst_mean_)
                                width_wst_median.append(width_wst_median_)
                                width_wst_min.append(width_wst_min_)
                                width_wst_max.append(width_wst_max_)

                                length_wst_mean.append(length_wst_mean_)
                                length_wst_median.append(length_wst_median_)
                                length_wst_min.append(length_wst_min_)
                                length_wst_max.append(length_wst_max_)

                                var_area_wst_mean.append(var_area_wst_mean_)
                                var_area_wst_median.append(var_area_wst_median_)
                                var_area_wst_min.append(var_area_wst_min_)
                                var_area_wst_max.append(var_area_wst_max_)

                                var_width_wst_mean.append(var_width_wst_mean_)
                                var_width_wst_median.append(var_width_wst_median_)
                                var_width_wst_min.append(var_width_wst_min_)
                                var_width_wst_max.append(var_width_wst_max_)

                                var_length_wst_mean.append(var_length_wst_mean_)
                                var_length_wst_median.append(var_length_wst_median_)
                                var_length_wst_min.append(var_length_wst_min_)
                                var_length_wst_max.append(var_length_wst_max_)    

                                ## for stomata
                                No_st_mean.append(No_st_mean_)
                                No_st_median.append(No_st_median_)
                                No_st_min.append(No_st_min_)
                                No_st_max.append(No_st_max_)

                                box_w_st_mean.append(box_w_st_mean_)
                                box_w_st_median.append(box_w_st_median_)
                                box_w_st_min.append(box_w_st_min_)
                                box_w_st_max.append(box_w_st_max_)

                                box_h_st_mean.append(box_h_st_mean_)
                                box_h_st_median.append(box_h_st_median_)
                                box_h_st_min.append(box_h_st_min_)
                                box_h_st_max.append(box_h_st_max_)

                                area_st_mean.append(area_st_mean_)
                                area_st_median.append(area_st_median_)
                                area_st_min.append(area_st_min_)
                                area_st_max.append(area_st_max_)

                                width_st_mean.append(width_st_mean_)
                                width_st_median.append(width_st_median_)
                                width_st_min.append(width_st_min_)
                                width_st_max.append(width_st_max_)

                                length_st_mean.append(length_st_mean_)
                                length_st_median.append(length_st_median_)
                                length_st_min.append(length_st_min_)
                                length_st_max.append(length_st_max_)

                                var_area_st_mean.append(var_area_st_mean_)
                                var_area_st_median.append(var_area_st_median_)
                                var_area_st_min.append(var_area_st_min_)
                                var_area_st_max.append(var_area_st_max_)

                                var_width_st_mean.append(var_width_st_mean_)
                                var_width_st_median.append(var_width_st_median_)
                                var_width_st_min.append(var_width_st_min_)
                                var_width_st_max.append(var_width_st_max_)

                                var_length_st_mean.append(var_length_st_mean_)
                                var_length_st_median.append(var_length_st_median_)
                                var_length_st_min.append(var_length_st_min_)
                                var_length_st_max.append(var_length_st_max_)  

                                ## for guard cell
                                guardCell_length_mean.append(guardCell_length_mean_) 
                                guardCell_length_median.append(guardCell_length_median_)  
                                guardCell_length_min.append(guardCell_length_min_)
                                guardCell_length_max.append(guardCell_length_max_)    

                                guardCell_width_mean.append(guardCell_width_mean_)
                                guardCell_width_median.append(guardCell_width_median_)
                                guardCell_width_min.append(guardCell_width_min_)
                                guardCell_width_max.append(guardCell_width_max_)

                                guardCell_area_mean.append(guardCell_area_mean_)
                                guardCell_area_median.append(guardCell_area_median_)
                                guardCell_area_min.append(guardCell_area_min_)
                                guardCell_area_max.append(guardCell_area_max_)

                                guardCell_angle_mean.append(guardCell_angle_mean_)
                                guardCell_angle_median.append(guardCell_angle_median_)
                                guardCell_angle_min.append(guardCell_angle_min_)
                                guardCell_angle_max.append(guardCell_angle_max_)

                                var_width_guardCell_mean.append(var_width_guardCell_mean_)
                                var_width_guardCell_median.append(var_width_guardCell_median_)
                                var_width_guardCell_min.append(var_width_guardCell_min_)
                                var_width_guardCell_max.append(var_width_guardCell_max_)

                                var_length_guardCell_mean.append(var_length_guardCell_mean_)
                                var_length_guardCell_median.append(var_length_guardCell_median_)
                                var_length_guardCell_min.append(var_length_guardCell_min_)
                                var_length_guardCell_max.append(var_length_guardCell_max_)

                                wst_density_mean.append(wst_density_mean_)
                                wst_density_median.append(wst_density_median_)
                                wst_density_min.append(wst_density_min_)
                                wst_density_max.append(wst_density_max_)

                                ratio_area_st_gc_mean.append(ratio_area_st_gc_mean_)
                                ratio_area_st_gc_median.append(ratio_area_st_gc_median_)
                                ratio_area_st_gc_min.append(ratio_area_st_gc_min_)
                                ratio_area_st_gc_max.append(ratio_area_st_gc_max_)

                                var_angle_mean.append(var_angle_mean_)
                                var_angle_median.append(var_angle_median_)
                                var_angle_min.append(var_angle_min_)
                                var_angle_max.append(var_angle_max_)

                                ratio_area_to_img_mean.append(ratio_area_to_img_mean_)
                                ratio_area_to_img_median.append(ratio_area_to_img_median_)
                                ratio_area_to_img_min.append(ratio_area_to_img_min_)
                                ratio_area_to_img_max.append(ratio_area_to_img_max_)

                                SEve_mean.append(SEve_mean_)
                                SEve_median.append(SEve_median_)
                                SEve_min.append(SEve_min_)
                                SEve_max.append(SEve_max_)

                                SDiv_mean.append(SDiv_mean_)
                                SDiv_median.append(SDiv_median_)
                                SDiv_min.append(SDiv_min_)
                                SDiv_max.append(SDiv_max_)

                                SAgg_mean.append(SAgg_mean_)
                                SAgg_median.append(SAgg_median_)
                                SAgg_min.append(SAgg_min_)
                                SAgg_max.append(SAgg_max_)

                                self.progressBar.setValue(int((self.img_num / len(glob.glob(output_image_path_ + '/' + '*.csv'))*100)))
                                QtWidgets.QApplication.processEvents()

                                self.img_num +=1 

                            except KeyError:
                                self.keyError()
                                Filename.pop()
                                self.img_num +=2                 

                    # Convert all lists into pd.series
                    Filename = pd.Series(Filename, dtype=pd.StringDtype(), name="Filename")
                    
                    No_wst_mean = pd.Series(No_wst_mean, dtype=pd.Float64Dtype(), name="No_wst_mean").map('{:,.0f}'.format)
                    No_wst_median = pd.Series(No_wst_median, dtype=pd.Float64Dtype(), name="No_wst_median").map('{:,.0f}'.format)
                    No_wst_min = pd.Series(No_wst_min, dtype=pd.Float64Dtype(), name="No_wst_min").map('{:,.0f}'.format)
                    No_wst_max = pd.Series(No_wst_max, dtype=pd.Float64Dtype(), name="No_wst_max").map('{:,.0f}'.format)

                    box_w_wst_mean = pd.Series(box_w_wst_mean, dtype=pd.Float64Dtype(), name="box_w_wst_mean").map('{:,.0f}'.format)
                    box_w_wst_median = pd.Series(box_w_wst_median, dtype=pd.Float64Dtype(), name="box_w_wst_median").map('{:,.0f}'.format)
                    box_w_wst_min = pd.Series(box_w_wst_min, dtype=pd.Float64Dtype(), name="box_w_wst_min").map('{:,.0f}'.format)
                    box_w_wst_max = pd.Series(box_w_wst_max, dtype=pd.Float64Dtype(), name="box_w_wst_max").map('{:,.0f}'.format)

                    box_h_wst_mean = pd.Series(box_h_wst_mean, dtype=pd.Float64Dtype(), name="box_h_wst_mean").map('{:,.0f}'.format)
                    box_h_wst_median = pd.Series(box_h_wst_median, dtype=pd.Float64Dtype(), name="box_h_wst_median").map('{:,.0f}'.format)
                    box_h_wst_min = pd.Series(box_h_wst_min, dtype=pd.Float64Dtype(), name="box_h_wst_min").map('{:,.0f}'.format)
                    box_h_wst_max = pd.Series(box_h_wst_max, dtype=pd.Float64Dtype(), name="box_h_wst_max").map('{:,.0f}'.format)

                    area_wst_mean = pd.Series(area_wst_mean, dtype=pd.Float64Dtype(), name="area_wst_mean").map('{:,.0f}'.format)
                    area_wst_median = pd.Series(area_wst_median, dtype=pd.Float64Dtype(), name="area_wst_median").map('{:,.0f}'.format)
                    area_wst_min = pd.Series(area_wst_min, dtype=pd.Float64Dtype(), name="area_wst_min").map('{:,.0f}'.format)
                    area_wst_max = pd.Series(area_wst_max, dtype=pd.Float64Dtype(), name="area_wst_max").map('{:,.0f}'.format)

                    width_wst_mean = pd.Series(width_wst_mean, dtype=pd.Float64Dtype(), name="width_wst_mean").map('{:,.0f}'.format)
                    width_wst_median = pd.Series(width_wst_median, dtype=pd.Float64Dtype(), name="width_wst_median").map('{:,.0f}'.format)
                    width_wst_min = pd.Series(width_wst_min, dtype=pd.Float64Dtype(), name="width_wst_min").map('{:,.0f}'.format)
                    width_wst_max = pd.Series(width_wst_max, dtype=pd.Float64Dtype(), name="width_wst_max").map('{:,.0f}'.format)

                    length_wst_mean = pd.Series(length_wst_mean, dtype=pd.Float64Dtype(), name="length_wst_mean").map('{:,.0f}'.format)
                    length_wst_median = pd.Series(length_wst_median, dtype=pd.Float64Dtype(), name="length_wst_median").map('{:,.0f}'.format)
                    length_wst_min = pd.Series(length_wst_min, dtype=pd.Float64Dtype(), name="length_wst_min").map('{:,.0f}'.format)
                    length_wst_max = pd.Series(length_wst_max, dtype=pd.Float64Dtype(), name="length_wst_max").map('{:,.0f}'.format)

                    var_area_wst_mean = pd.Series(var_area_wst_mean, dtype=pd.Float64Dtype(), name="var_area_wst_mean").map('{:,.0f}'.format)
                    var_area_wst_median = pd.Series(var_area_wst_median, dtype=pd.Float64Dtype(), name="var_area_wst_median").map('{:,.0f}'.format)
                    var_area_wst_min = pd.Series(var_area_wst_min, dtype=pd.Float64Dtype(), name="var_area_wst_min").map('{:,.0f}'.format)
                    var_area_wst_max = pd.Series(var_area_wst_max, dtype=pd.Float64Dtype(), name="var_area_wst_max").map('{:,.0f}'.format)

                    var_width_wst_mean = pd.Series(var_width_wst_mean, dtype=pd.Float64Dtype(), name="var_width_wst_mean").map('{:,.0f}'.format)
                    var_width_wst_median = pd.Series(var_width_wst_median, dtype=pd.Float64Dtype(), name="var_width_wst_median").map('{:,.0f}'.format)
                    var_width_wst_min = pd.Series(var_width_wst_min, dtype=pd.Float64Dtype(), name="var_width_wst_min").map('{:,.0f}'.format)
                    var_width_wst_max = pd.Series(var_width_wst_max, dtype=pd.Float64Dtype(), name="var_width_wst_max").map('{:,.0f}'.format)       

                    var_length_wst_mean = pd.Series(var_length_wst_mean, dtype=pd.Float64Dtype(), name="var_length_wst_mean").map('{:,.0f}'.format)
                    var_length_wst_median = pd.Series(var_length_wst_median, dtype=pd.Float64Dtype(), name="var_length_wst_median").map('{:,.0f}'.format)
                    var_length_wst_min = pd.Series(var_length_wst_min, dtype=pd.Float64Dtype(), name="var_length_wst_min").map('{:,.0f}'.format)
                    var_length_wst_max = pd.Series(var_length_wst_max, dtype=pd.Float64Dtype(), name="var_length_wst_max").map('{:,.0f}'.format)   

                    No_st_mean = pd.Series(No_st_mean, dtype=pd.Float64Dtype(), name="No_st_mean").map('{:,.0f}'.format)
                    No_st_median = pd.Series(No_st_median, dtype=pd.Float64Dtype(), name="No_st_median").map('{:,.0f}'.format)
                    No_st_min = pd.Series(No_st_min, dtype=pd.Float64Dtype(), name="No_st_min").map('{:,.0f}'.format)
                    No_st_max = pd.Series(No_st_max, dtype=pd.Float64Dtype(), name="No_st_max").map('{:,.0f}'.format)   

                    box_w_st_mean = pd.Series(box_w_st_mean, dtype=pd.Float64Dtype(), name="box_w_st_mean").map('{:,.0f}'.format)
                    box_w_st_median = pd.Series(box_w_st_median, dtype=pd.Float64Dtype(), name="box_w_st_median").map('{:,.0f}'.format)
                    box_w_st_min = pd.Series(box_w_st_min, dtype=pd.Float64Dtype(), name="box_w_st_min").map('{:,.0f}'.format)
                    box_w_st_max = pd.Series(box_w_st_max, dtype=pd.Float64Dtype(), name="box_w_st_max").map('{:,.0f}'.format)   

                    box_h_st_mean = pd.Series(box_h_st_mean, dtype=pd.Float64Dtype(), name="box_h_st_mean").map('{:,.0f}'.format)
                    box_h_st_median = pd.Series(box_h_st_median, dtype=pd.Float64Dtype(), name="box_h_st_median").map('{:,.0f}'.format)
                    box_h_st_min = pd.Series(box_h_st_min, dtype=pd.Float64Dtype(), name="box_h_st_min").map('{:,.0f}'.format)
                    box_h_st_max = pd.Series(box_h_st_max, dtype=pd.Float64Dtype(), name="box_h_st_max").map('{:,.0f}'.format)   

                    area_st_mean = pd.Series(area_st_mean, dtype=pd.Float64Dtype(), name="area_st_mean").map('{:,.0f}'.format)
                    area_st_median = pd.Series(area_st_median, dtype=pd.Float64Dtype(), name="area_st_median").map('{:,.0f}'.format)
                    area_st_min = pd.Series(area_st_min, dtype=pd.Float64Dtype(), name="area_st_min").map('{:,.0f}'.format)
                    area_st_max = pd.Series(area_st_max, dtype=pd.Float64Dtype(), name="area_st_max").map('{:,.0f}'.format)   

                    width_st_mean = pd.Series(width_st_mean, dtype=pd.Float64Dtype(), name="width_st_mean").map('{:,.0f}'.format)
                    width_st_median = pd.Series(width_st_median, dtype=pd.Float64Dtype(), name="width_st_median").map('{:,.0f}'.format)
                    width_st_min = pd.Series(width_st_min, dtype=pd.Float64Dtype(), name="width_st_min").map('{:,.0f}'.format)
                    width_st_max = pd.Series(width_st_max, dtype=pd.Float64Dtype(), name="width_st_max").map('{:,.0f}'.format)  

                    length_st_mean = pd.Series(length_st_mean, dtype=pd.Float64Dtype(), name="length_st_mean").map('{:,.0f}'.format)
                    length_st_median = pd.Series(length_st_median, dtype=pd.Float64Dtype(), name="length_st_median").map('{:,.0f}'.format)
                    length_st_min = pd.Series(length_st_min, dtype=pd.Float64Dtype(), name="length_st_min").map('{:,.0f}'.format)
                    length_st_max = pd.Series(length_st_max, dtype=pd.Float64Dtype(), name="length_st_max").map('{:,.0f}'.format)   

                    var_area_st_mean = pd.Series(var_area_st_mean, dtype=pd.Float64Dtype(), name="var_area_st_mean").map('{:,.0f}'.format)
                    var_area_st_median = pd.Series(var_area_st_median, dtype=pd.Float64Dtype(), name="var_area_st_median").map('{:,.0f}'.format)
                    var_area_st_min = pd.Series(var_area_st_min, dtype=pd.Float64Dtype(), name="var_area_st_min").map('{:,.0f}'.format)
                    var_area_st_max = pd.Series(var_area_st_max, dtype=pd.Float64Dtype(), name="var_area_st_max").map('{:,.0f}'.format)   

                    var_width_st_mean = pd.Series(var_width_st_mean, dtype=pd.Float64Dtype(), name="var_width_st_mean").map('{:,.0f}'.format)
                    var_width_st_median = pd.Series(var_width_st_median, dtype=pd.Float64Dtype(), name="var_width_st_median").map('{:,.0f}'.format)
                    var_width_st_min = pd.Series(var_width_st_min, dtype=pd.Float64Dtype(), name="var_width_st_min").map('{:,.0f}'.format)
                    var_width_st_max = pd.Series(var_width_st_max, dtype=pd.Float64Dtype(), name="var_width_st_max").map('{:,.0f}'.format)   

                    var_length_st_mean = pd.Series(var_length_st_mean, dtype=pd.Float64Dtype(), name="var_length_st_mean").map('{:,.0f}'.format)
                    var_length_st_median = pd.Series(var_length_st_median, dtype=pd.Float64Dtype(), name="var_length_st_median").map('{:,.0f}'.format)
                    var_length_st_min = pd.Series(var_length_st_min, dtype=pd.Float64Dtype(), name="var_length_st_min").map('{:,.0f}'.format)
                    var_length_st_max = pd.Series(var_length_st_max, dtype=pd.Float64Dtype(), name="var_length_st_max").map('{:,.0f}'.format)   

                    guardCell_length_mean = pd.Series(guardCell_length_mean, dtype=pd.Float64Dtype(), name="guardCell_length_mean").map('{:,.0f}'.format)
                    guardCell_length_median = pd.Series(guardCell_length_median, dtype=pd.Float64Dtype(), name="guardCell_length_median").map('{:,.0f}'.format)
                    guardCell_length_min = pd.Series(guardCell_length_min, dtype=pd.Float64Dtype(), name="guardCell_length_min").map('{:,.0f}'.format)
                    guardCell_length_max = pd.Series(guardCell_length_max, dtype=pd.Float64Dtype(), name="guardCell_length_max").map('{:,.0f}'.format)   

                    guardCell_width_mean = pd.Series(guardCell_width_mean, dtype=pd.Float64Dtype(), name="guardCell_width_mean").map('{:,.0f}'.format)
                    guardCell_width_median = pd.Series(guardCell_width_median, dtype=pd.Float64Dtype(), name="guardCell_width_median").map('{:,.0f}'.format)
                    guardCell_width_min = pd.Series(guardCell_width_min, dtype=pd.Float64Dtype(), name="guardCell_width_min").map('{:,.0f}'.format)
                    guardCell_width_max = pd.Series(guardCell_width_max, dtype=pd.Float64Dtype(), name="guardCell_width_max").map('{:,.0f}'.format)  

                    guardCell_area_mean = pd.Series(guardCell_area_mean, dtype=pd.Float64Dtype(), name="guardCell_area_mean").map('{:,.0f}'.format)
                    guardCell_area_median = pd.Series(guardCell_area_median, dtype=pd.Float64Dtype(), name="guardCell_area_median").map('{:,.0f}'.format)
                    guardCell_area_min = pd.Series(guardCell_area_min, dtype=pd.Float64Dtype(), name="guardCell_area_min").map('{:,.0f}'.format)
                    guardCell_area_max = pd.Series(guardCell_area_max, dtype=pd.Float64Dtype(), name="guardCell_area_max").map('{:,.0f}'.format) 

                    guardCell_angle_mean = pd.Series(guardCell_angle_mean, dtype=pd.Float64Dtype(), name="guardCell_angle_mean").map('{:,.0f}'.format)
                    guardCell_angle_median = pd.Series(guardCell_angle_median, dtype=pd.Float64Dtype(), name="guardCell_angle_median").map('{:,.0f}'.format)
                    guardCell_angle_min = pd.Series(guardCell_angle_min, dtype=pd.Float64Dtype(), name="guardCell_angle_min").map('{:,.0f}'.format)
                    guardCell_angle_max = pd.Series(guardCell_angle_max, dtype=pd.Float64Dtype(), name="guardCell_angle_max").map('{:,.0f}'.format)  

                    var_width_guardCell_mean = pd.Series(var_width_guardCell_mean, dtype=pd.Float64Dtype(), name="var_width_guardCell_mean").map('{:,.0f}'.format)
                    var_width_guardCell_median = pd.Series(var_width_guardCell_median, dtype=pd.Float64Dtype(), name="var_width_guardCell_median").map('{:,.0f}'.format)
                    var_width_guardCell_min = pd.Series(var_width_guardCell_min, dtype=pd.Float64Dtype(), name="var_width_guardCell_min").map('{:,.0f}'.format)
                    var_width_guardCell_max = pd.Series(var_width_guardCell_max, dtype=pd.Float64Dtype(), name="var_width_guardCell_max").map('{:,.0f}'.format)  

                    var_length_guardCell_mean = pd.Series(var_length_guardCell_mean, dtype=pd.Float64Dtype(), name="var_length_guardCell_mean").map('{:,.0f}'.format)
                    var_length_guardCell_median = pd.Series(var_length_guardCell_median, dtype=pd.Float64Dtype(), name="var_length_guardCell_median").map('{:,.0f}'.format)
                    var_length_guardCell_min = pd.Series(var_length_guardCell_min, dtype=pd.Float64Dtype(), name="var_length_guardCell_min").map('{:,.0f}'.format)
                    var_length_guardCell_max = pd.Series(var_length_guardCell_max, dtype=pd.Float64Dtype(), name="var_length_guardCell_max").map('{:,.0f}'.format)  

                    wst_density_mean = pd.Series(wst_density_mean, dtype=pd.Float64Dtype(), name="wst_density_mean").map('{:,.0f}'.format)
                    wst_density_median = pd.Series(wst_density_median, dtype=pd.Float64Dtype(), name="wst_density_median").map('{:,.0f}'.format)
                    wst_density_min = pd.Series(wst_density_min, dtype=pd.Float64Dtype(), name="wst_density_min").map('{:,.0f}'.format)
                    wst_density_max = pd.Series(wst_density_max, dtype=pd.Float64Dtype(), name="wst_density_max").map('{:,.0f}'.format)  

                    ratio_area_st_gc_mean = pd.Series(ratio_area_st_gc_mean, dtype=pd.Float64Dtype(), name="ratio_area_st_gc_mean").map('{:,.3f}'.format)
                    ratio_area_st_gc_median = pd.Series(ratio_area_st_gc_median, dtype=pd.Float64Dtype(), name="ratio_area_st_gc_median").map('{:,.3f}'.format)
                    ratio_area_st_gc_min = pd.Series(ratio_area_st_gc_min, dtype=pd.Float64Dtype(), name="ratio_area_st_gc_min").map('{:,.3f}'.format)
                    ratio_area_st_gc_max = pd.Series(ratio_area_st_gc_max, dtype=pd.Float64Dtype(), name="ratio_area_st_gc_max").map('{:,.3f}'.format) 

                    ratio_area_to_img_mean = pd.Series(ratio_area_to_img_mean, dtype=pd.Float64Dtype(), name="ratio_area_to_img_mean").map('{:,.3f}'.format)
                    ratio_area_to_img_median = pd.Series(ratio_area_to_img_median, dtype=pd.Float64Dtype(), name="ratio_area_to_img_median").map('{:,.3f}'.format)
                    ratio_area_to_img_min = pd.Series(ratio_area_to_img_min, dtype=pd.Float64Dtype(), name="ratio_area_to_img_min").map('{:,.3f}'.format)
                    ratio_area_to_img_max = pd.Series(ratio_area_to_img_max, dtype=pd.Float64Dtype(), name="ratio_area_to_img_max").map('{:,.3f}'.format) 

                    var_angle_mean = pd.Series(var_angle_mean, dtype=pd.Float64Dtype(), name="var_angle_mean").map('{:,.0f}'.format)
                    var_angle_median = pd.Series(var_angle_median, dtype=pd.Float64Dtype(), name="var_angle_median").map('{:,.0f}'.format)
                    var_angle_min = pd.Series(var_angle_min, dtype=pd.Float64Dtype(), name="var_angle_min").map('{:,.0f}'.format)
                    var_angle_max = pd.Series(var_angle_max, dtype=pd.Float64Dtype(), name="var_angle_max").map('{:,.0f}'.format) 

                    SEve_mean = pd.Series(SEve_mean, dtype=pd.Float64Dtype(), name="SEve_mean").map('{:,.4f}'.format)
                    SEve_median = pd.Series(SEve_median, dtype=pd.Float64Dtype(), name="SEve_median").map('{:,.4f}'.format)
                    SEve_min = pd.Series(SEve_min, dtype=pd.Float64Dtype(), name="SEve_min").map('{:,.4f}'.format)
                    SEve_max = pd.Series(SEve_max, dtype=pd.Float64Dtype(), name="SEve_max").map('{:,.4f}'.format) 

                    SDiv_mean = pd.Series(SDiv_mean, dtype=pd.Float64Dtype(), name="SDiv_mean").map('{:,.4f}'.format)
                    SDiv_median = pd.Series(SDiv_median, dtype=pd.Float64Dtype(), name="SDiv_median").map('{:,.4f}'.format)
                    SDiv_min = pd.Series(SDiv_min, dtype=pd.Float64Dtype(), name="SDiv_min").map('{:,.4f}'.format)
                    SDiv_max = pd.Series(SDiv_max, dtype=pd.Float64Dtype(), name="SDiv_max").map('{:,.4f}'.format) 

                    SAgg_mean = pd.Series(SAgg_mean, dtype=pd.Float64Dtype(), name="SAgg_mean").map('{:,.4f}'.format)
                    SAgg_median = pd.Series(SAgg_median, dtype=pd.Float64Dtype(), name="SAgg_median").map('{:,.4f}'.format)
                    SAgg_min = pd.Series(SAgg_min, dtype=pd.Float64Dtype(), name="SAgg_min").map('{:,.4f}'.format)
                    SAgg_max = pd.Series(SAgg_max, dtype=pd.Float64Dtype(), name="SAgg_max").map('{:,.4f}'.format)                      

                    Site = pd.Series(Site, dtype=pd.StringDtype(), name="Site")
                    Block = pd.Series(Block, dtype=pd.StringDtype(), name="Block")
                    Clone = pd.Series(Clone, dtype=pd.StringDtype(), name="Clone")
                    Month = pd.Series(Month, dtype=pd.StringDtype(), name="Month")
                    Year = pd.Series(Year, dtype=pd.StringDtype(), name="Year")
                    # Put all extracted parameters into a data frame
                    Output_data = pd.concat(
                        [Site, 
                        Block, 
                        Clone, 
                        Month, 
                        Year,
                        Filename, 
                        No_wst_mean, 
                        No_wst_median,
                        No_wst_min, 
                        No_wst_max, 
                        box_w_wst_mean, 
                        box_w_wst_median,
                        box_w_wst_min, 
                        box_w_wst_max, 
                        box_h_wst_mean, 
                        box_h_wst_median, 
                        box_h_wst_min,
                        box_h_wst_max,
                        area_wst_mean,
                        area_wst_median,
                        area_wst_min,
                        area_wst_max,
                        width_wst_mean,
                        width_wst_median,
                        width_wst_min,
                        width_wst_max,
                        length_wst_mean,
                        length_wst_median,
                        length_wst_min,
                        length_wst_max,
                        var_area_wst_mean,
                        var_area_wst_median,
                        var_area_wst_min,
                        var_area_wst_max,
                        var_width_wst_mean,
                        var_width_wst_median,
                        var_width_wst_min,
                        var_width_wst_max,
                        var_length_wst_mean,
                        var_length_wst_median,
                        var_length_wst_min,
                        var_length_wst_max,
                        No_st_mean,
                        No_st_median,
                        No_st_min,
                        No_st_max,
                        box_w_st_mean,
                        box_w_st_median,
                        box_w_st_min,
                        box_w_st_max,
                        box_h_st_mean,
                        box_h_st_median,
                        box_h_st_min,
                        box_h_st_max,
                        area_st_mean,
                        area_st_median,
                        area_st_min,
                        area_st_max,
                        width_st_mean,
                        width_st_median,
                        width_st_min,
                        width_st_max,
                        length_st_mean,
                        length_st_median,
                        length_st_min,
                        length_st_max,
                        var_area_st_mean,
                        var_area_st_median,
                        var_area_st_min,
                        var_area_st_max,
                        var_width_st_mean,
                        var_width_st_median,
                        var_width_st_min,
                        var_width_st_max,
                        var_length_st_mean,
                        var_length_st_median,
                        var_length_st_min,
                        var_length_st_max,
                        guardCell_length_mean,
                        guardCell_length_median,
                        guardCell_length_min,
                        guardCell_length_max,
                        guardCell_width_mean,
                        guardCell_width_median,
                        guardCell_width_min,
                        guardCell_width_max,
                        guardCell_area_mean,
                        guardCell_area_median,
                        guardCell_area_min,
                        guardCell_area_max,
                        guardCell_angle_mean,
                        guardCell_angle_median,
                        guardCell_angle_min,
                        guardCell_angle_max,
                        var_width_guardCell_mean,
                        var_width_guardCell_median,
                        var_width_guardCell_min,
                        var_width_guardCell_max,
                        var_length_guardCell_mean,
                        var_length_guardCell_median,
                        var_length_guardCell_min,
                        var_length_guardCell_max,
                        wst_density_mean,
                        wst_density_median,
                        wst_density_min,
                        wst_density_max,
                        ratio_area_st_gc_mean,
                        ratio_area_st_gc_median,
                        ratio_area_st_gc_min,
                        ratio_area_st_gc_max,
                        ratio_area_to_img_mean,
                        ratio_area_to_img_median,
                        ratio_area_to_img_min,
                        ratio_area_to_img_max,
                        var_angle_mean,
                        var_angle_median,
                        var_angle_min,
                        var_angle_max,
                        SEve_mean,
                        SEve_median,
                        SEve_min,
                        SEve_max,
                        SDiv_mean,
                        SDiv_median,
                        SDiv_min,
                        SDiv_max,
                        SAgg_mean,
                        SAgg_median,
                        SAgg_min,
                        SAgg_max], 
                        axis=1)
                    
                    # check if the files are currently opening ? if yes, create a copy, if not replace the older ones
                    try: 
                        Output_data = pd.DataFrame(data=Output_data)
                        random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
                        Output_data.to_excel(output_image_path_ + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
                        self.show_info_messagebox_group_analysis()

                    except OSError:
                        self.closeExcel()                    

                else:
                    self.show_info_messagebox_group_analysis_2()
                    pass         

    def Stomata_no_groups_analysis(self):
        """ """
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
                    
                    if single_csv_file.shape[0]>=4:

                        try:
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
                            Whole_stomata_areas_max = max(Whole_stomata_areas)
                            Whole_stomata_areas_min = min(Whole_stomata_areas)
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
                            WST_Density.append(int(np.mean(Whole_stomata_density)))
                            WST_Number.append(np.mean(Whole_stomata_number))
                            WST_Area.append(np.mean(Whole_stomata_areas))
                            WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))
                            
                            self.progressBar.setValue(int((self.img_num / len(glob.glob(self.output_image_path_ + '/' + '*.csv'))*100)))
                            QtWidgets.QApplication.processEvents()

                            self.img_num +=1

                        except KeyError:
                            self.keyError()
                            self.img_num +=2                        

                    else:
                        self.img_num +=1
                        pass                                        

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

    # define a function use YOLOv8-seg-x model 

    def Check_input_path_folder(self):
        self.folder = self.lineEdit.text()
        list_dir = os.listdir(self.folder)
        list_folder = bool()
        for f in list_dir:
            if not os.path.isfile(os.path.join(self.folder, f)):
                list_folder=True
                break
            else:
                list_folder = False
        if list_folder is True:
            self.messagebox_if_input_path_contains_folder()
        else: 
            self.GuardCell()

    def GuardCell(self):
        """ the main function to segment and analyze stomatal and guard cell metrics """

        self.folder = self.lineEdit.text()
        self.output_folder = self.lineEdit_2.text()
        self.p = self.lineEdit_4.text()
        self.confidence = self.lineEdit_5.text()
        filename_without_ext = str()

        # Check for image file extensions
        image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
        has_image_files = any(glob.glob(os.path.join(self.input_image_path_, f"*{ext}")) for ext in image_extensions)

        if has_image_files:
            # To check if the input folder is set
            # Create output path for saving output imgs and csv files
            print(self.output_folder)

            if not os.path.exists(str(os.path.join(str(self.output_folder), "Predict_output"))):
                os.makedirs(str(os.path.join(str(self.output_folder), "Predict_output")))
                output_path = str(os.path.join(str(self.output_folder), "Predict_output"))
            else:
                output_path = str(os.path.join(str(self.output_folder), "Predict_output"))

            self.img_num = 1
            for filename in os.listdir(self.folder):
                try:    
                    imgPath = os.path.join(self.folder,filename)

                    torch.cuda.empty_cache()

                    self.progressBar.setValue(int((self.img_num / len(os.listdir(self.folder))*100)))
                    QtWidgets.QApplication.processEvents()

                    self.img_num +=1

                    # Remove the file extension ".jpg"
                    filename_without_ext = filename.replace(".jpg",'')  

                    if os.path.exists(str(os.path.join(str(output_path), "Output_csv"))) is False:
                        os.makedirs(str(os.path.join(str(output_path), "Output_csv")))
                        new_path = str(os.path.join(str(output_path), "Output_csv"))
                    else:
                        new_path = str(os.path.join(str(output_path), "Output_csv"))
                        pass

                    ## define a function to calculate guard cell infomation
                    def StomGuardCell(imgPath, p):
                        """ do prediction and extract import features """

                        images = cv2.imread(imgPath) # Define the image for predicting
                
                        ori_img_shape = (images.shape[1],images.shape[0])
                        if  images.shape[1]>1280: # reduce the image size if the image is too large, which could take up all GPU memory
                            scale_percent = 50 # percent of original size
                            width = int(images.shape[1] * scale_percent / 100)
                            height = int(images.shape[0] * scale_percent / 100)
                            dim = (width, height)

                            images = cv2.resize(images, dim, interpolation = cv2.INTER_AREA)
                        else:
                            scale_percent = 100 # percent of original size

                        self.confidence = self.lineEdit_5.text()

                        self.p = self.lineEdit_4.text()

                        if self.p =="Input pixels in 0.1 mm, the default is: 465" or self.p =="" or self.p ==" ":
                            self.p = 465
                        else: 
                            self.p = int(self.lineEdit_4.text())

                        if self.confidence =="Confidence threshold for detection the default is 0.25" or self.confidence =="" or self.confidence ==" ":
                            self.confidence = 0.25
                        else: 
                            self.confidence = float(self.lineEdit_5.text())

                        model = YOLO("best.pt") # config the model

                        results = model.predict(images, save=True, save_txt = True, imgsz=640, conf=self.confidence,  
                                                save_crop=False, line_thickness = 1, retina_masks = True, project = str(new_path), name = f""+filename_without_ext,
                                                max_det=500) # implement predicting and pass the output to results
                        
                        H, W, _ = images.shape # obtain the image height and width for denormalization (get absolute pixels)
                        stomata = {"class":[], "box_w":[], "box_h":[], "area":[],"centorid":[],"length":[],"width":[],"angle":[],"img_shape":[]} # create a stomata dictionary
                        whole_stomata = {"class":[],"box_w":[], "box_h":[],"area":[],"centorid":[],"length":[],"width":[],"angle":[],"img_shape":[]} # create a whole_stomata dictionary

                        def PolyArea(x,y): 
                            """ to calculate the mask segments area """
                            
                            return 0.5*np.abs(np.dot(x,np.roll(y,1))-np.dot(y,np.roll(x,1))) 
                            
                        i = 0
                        for i in range(len(results[0].masks.xyn)): # build a for loop to extract segment elements

                            box = results[0].boxes[i]
                            box_w = box.xywh.tolist()[0][2]*(100/scale_percent) # the bounding box with xywh format, (N, 4)
                            box_h = box.xywh.tolist()[0][3]*(100/scale_percent) # the bounding box with xywh format, (N, 4)

                            # boxes.xyxy  # box with xyxy format, (N, 4)
                            # boxes.xywh  # box with xywh format, (N, 4)
                            # boxes.xyxyn  # box with xyxy format but normalized, (N, 4)
                            # boxes.xywhn  # box with xywh format but normalized, (N, 4)
                            # boxes.conf  # confidence score, (N, 1)
                            # boxes.cls  # cls, (N, 1)
                            # boxes.data  # raw bboxes tensor, (N, 6) or boxes.boxes

                            cls = str(int(box.cls.tolist()[0]))

                            x = (results[0].masks.xyn[i][:,0]*W).astype("int")
                            y = (results[0].masks.xyn[i][:,1]*H).astype("int")

                            # create example polygon
                            poly = Polygon(zip(x,y))
                            # get minimum bounding box around polygon
                            box_ = poly.minimum_rotated_rectangle
                            # get coordinates of polygon vertices

                            x1, y1 = box_.exterior.coords.xy

                            # get length of bounding box edges
                            edge_length = ((Point(x1[0], y1[0]).distance(Point(x1[1], y1[1])))*(100/scale_percent), (Point(x1[1], y1[1]).distance(Point(x1[2], y1[2])))*(100/scale_percent))
                            # get length of polygon as the longest edge of the bounding box
                            length = int(max(edge_length))
                            # get width of polygon as the shortest edge of the bounding box
                            width = int(min(edge_length)) 
                            seg_area = int(PolyArea(x,y))*((100/scale_percent)*(100/scale_percent))
                            centroid = (int(sum(x) / len(x)), int(sum(y) / len(y)))
                        
                            def _azimuth(point1, point2):
                                """azimuth between 2 points (interval 0 - 180)"""
                                angle = np.arctan2(point2[0] - point1[0], point2[1] - point1[1])
                                return np.degrees(angle)  if angle > 0 else np.degrees(angle) + 180

                            def _dist(a, b):
                                """distance between points"""
                                return math.hypot(b[0] - a[0], b[1] - a[1])

                            def azimuth(mrr):
                                """azimuth of minimum_rotated_rectangle"""
                                bbox = list(mrr.exterior.coords)
                                axis1 = _dist(bbox[0], bbox[3])
                                axis2 = _dist(bbox[0], bbox[1])

                                if axis1 <= axis2:
                                    az = _azimuth(bbox[0], bbox[1])
                                else:
                                    az = _azimuth(bbox[0], bbox[3])
                                return az

                            angle = int(azimuth(box_))       
                        
                            if cls=="1":
                                whole_stomata["class"].append(cls)
                                whole_stomata["box_w"].append(box_w)
                                whole_stomata["box_h"].append(box_h)
                                whole_stomata["area"].append(seg_area)
                                whole_stomata["centorid"].append(centroid)
                                whole_stomata["length"].append(length) 
                                whole_stomata["width"].append(width)
                                whole_stomata["angle"].append(angle)
                                whole_stomata["img_shape"].append(ori_img_shape)      
                                i +=1
                            else:
                                stomata["class"].append(cls)
                                stomata["box_w"].append(box_w)
                                stomata["box_h"].append(box_h)
                                stomata["area"].append(seg_area)
                                stomata["centorid"].append(centroid)
                                stomata["length"].append(length) 
                                stomata["width"].append(width)
                                stomata["angle"].append(angle)
                                stomata["img_shape"].append(ori_img_shape)                 
                                i +=1
                                
                        res_whole_stomata = list(map(list, (whole_stomata.values()))) # convert the dictionary to multiple lists
                        res_stomata = list(map(list, (stomata.values()))) # convert the dictionary to multiple lists
                        
                        link_st_wst = [] # the list to combine all extracted data
                        guard_cell_area = [] # the list to collect all guard cell area
                        guard_cell_length = [] # the list to collect all guard cell length 
                        guard_cell_width =[] # the list to collect all guard cell width
                        guard_cell_angle = [] # the list to collect all guard cell angle
                        box_w_all = [] # the list to collect all bounding boxes w
                        box_h_all = [] # the list to collect all bounding boxes h
                        aperture_width = []
                        area_wst = []
                        area_st = []  
                        width_wst = []
                        width_st = []
                        length_wst = []
                        length_st = []            

                        j=0

                        # The following codes are used to calculate stomatal indices such as stomata_evenness, sum_deviance, stomata_aggregation
                        whole_stomata_centroid_all = res_whole_stomata[4]
                        stomata_num = (np.array(whole_stomata_centroid_all).shape[0])
                        dist_matrix = distance_matrix(whole_stomata_centroid_all, whole_stomata_centroid_all)
                        mst = minimum_spanning_tree(dist_matrix).toarray()
                        PD = np.sum(mst, axis=1) / np.sum(mst)
                        constant = 1 / (stomata_num - 1)
                        stomata_evenness = (np.sum(PD[PD < constant]) + (stomata_num - 1 - len(PD[PD < constant])) * constant - constant) / (1 - constant)                    
                        distance_to_gravity = np.array([distance.euclidean(whole_stomata_centroid_all[i], np.mean(whole_stomata_centroid_all, axis=0)) for i in range(stomata_num)])
                        mean_distance = np.mean(distance_to_gravity)                    
                        sum_deviance = np.sum(distance_to_gravity - mean_distance)
                        sum_abs_deviance = np.sum(np.abs(distance_to_gravity - mean_distance))                    
                        stomata_divergence = (sum_deviance + mean_distance) / (sum_abs_deviance + mean_distance)                    
                        stomatal_density = stomata_num / (ori_img_shape[0]*ori_img_shape[1]) # this is based on pixel level
                        theoretical_distance = 1 / (2 * (stomatal_density ** 0.5)) 
                        nearest_neighbor_distance = np.array([np.sort(dist_matrix[i])[1] for i in range(stomata_num)])
                        observed_distance = np.sum(nearest_neighbor_distance) / stomata_num
                        stomata_aggregation = observed_distance / theoretical_distance

                        for j in range(len(res_whole_stomata[0])): # Build a for loop to match stomata with whole_stomata to calculate guard cell width, length, and area
                            
                            class_1 = res_whole_stomata[0][j]
                            box_w_1 = res_whole_stomata[1][j]
                            box_h_1 = res_whole_stomata[2][j]
                            area_1 = int(((res_whole_stomata[3][j])/(100*self.p*self.p))*1000000)
                            
                            whole_stomata_centroid = res_whole_stomata[4][j]
                            length_1 = res_whole_stomata[5][j]
                            width_1 = res_whole_stomata[6][j]
                            angle_1 = res_whole_stomata[7][j]
                            k = 0

                            for k in range(len(res_stomata[0])):
                                
                                class_0 = res_stomata[0][k]
                                box_w_0 = res_stomata[1][k]
                                box_h_0 = res_stomata[2][k]
                                area_0 = int(((res_stomata[3][k])/(100*self.p*self.p))*1000000)
                                
                                stomata_centroid = res_stomata[4][k]
                                length_0 = res_stomata[5][k]
                                width_0 = res_stomata[6][k]
                                angle_0 = res_stomata[7][k]
                                
                                if math.hypot(stomata_centroid[0] - whole_stomata_centroid[0],
                                            stomata_centroid[1] - whole_stomata_centroid[1])<=30 and length_1>length_0 and area_1-area_0>0 and width_1-width_0>0: 
                                    # find the stomata within the whole_stomata and link them together
                                    guard_cell_area_ = area_1-area_0
                                    guard_cell_length_ = length_1/(self.p/100)
                                    aperture_width_ = width_0/(self.p/100)
                                    box_w_ = box_w/(self.p/100)
                                    box_h_ = box_h/(self.p/100)
                                    
                                    guard_cell_width_ = 0.5*((width_1-width_0)/(self.p/100))
                                    guard_cell_width.append(guard_cell_width_)
                                    guard_cell_length.append(guard_cell_length_)
                                    guard_cell_area.append(guard_cell_area_)
                                    guard_cell_angle.append(angle_1)
                                    aperture_width.append(aperture_width_)
                                    box_w_all.append(box_w_)
                                    box_h_all.append(box_h_)
                                    area_wst.append(area_1)
                                    area_st.append(area_0)
                                    width_wst.append(width_1/(self.p/100))
                                    width_st.append(width_0/(self.p/100))
                                    length_wst.append(length_1/(self.p/100))
                                    length_st.append(length_0/(self.p/100))                       

                                    var_area_wst = int(np.var(area_wst))
                                    var_area_st = int(np.var(area_st))
                                    var_angle = int(np.var(guard_cell_angle))
                                    var_width_wst = int(np.var(width_wst))
                                    var_width_st = int(np.var(width_st))
                                    var_length_wst = int(np.var(length_wst))
                                    var_length_st = int(np.var(length_st))
                                    var_width_guardCell = int(np.var(guard_cell_width))
                                    var_length_guardCell = int(np.var(guard_cell_length))
                                    var_area_guardCell = int(np.var(guard_cell_area))

                                    number_st = results[0].boxes.cls.tolist().count(0.0)
                                    number_wst = results[0].boxes.cls.tolist().count(1.0)
                                    img_area = ori_img_shape[0]*ori_img_shape[1]
                                    wst_density = int((number_wst)*((100*self.p*self.p)/img_area)) ## the unit is number of stomatal per 1 mm^2
                                    if area_0 == 0 or guard_cell_area_ ==0:
                                        ratio_area_st_gc = 0
                                    else:
                                        ratio_area_st_gc = float("{:.3f}".format(area_0/guard_cell_area_))   

                                    ratio_area_to_img = float("{:.3f}".format(sum((area/(10000/(self.p*self.p))) for area in area_wst) / img_area)) 

                                    
                                    link_st_wst.append([ori_img_shape, class_1, number_wst, j, int(box_w_1), int(box_h_1), 
                                                        area_1, int(width_1/(self.p/100)), int(length_1/(self.p/100)) ,var_area_wst,var_width_wst,var_length_wst, whole_stomata_centroid, class_0, number_st, 
                                                        k, int(box_w_0), int(box_h_0), area_0, int(width_0/(self.p/100)), 
                                                        int(length_0/(self.p/100)), var_area_st,var_width_st,var_length_st,stomata_centroid, guard_cell_length_,
                                                        guard_cell_width_, guard_cell_area_, angle_0, var_angle,var_width_guardCell, var_length_guardCell,var_area_guardCell, wst_density,ratio_area_st_gc, ratio_area_to_img,
                                                        float("{:.4f}".format(stomata_evenness)), float("{:.4f}".format(stomata_divergence)), float("{:.4f}".format(stomata_aggregation))])
                                    k+=1
                                else:
                                    pass
                                    k+=1
                            j +=1

                        df = pd.DataFrame(link_st_wst, columns = ['ori_img_shape', 'class_wst', 'number_wst', 'index_wst', 'box_w_wst', 'box_h_wst', 
                                                        'area_wst', 'width_wst', 'length_wst' ,'var_area_wst','var_width_wst','var_length_wst', 
                                                        'centroid_wst', 'class_st', 'number_st', 
                                                        'index_st', 'box_w_st', 'box_h_st', 'area_st', 'width_st', 
                                                        'length_st', 'var_area_st','var_width_st','var_length_st','centroid_st','guardCell_length',
                                                        'guardCell_width', 'guardCell_area', 'guardCell_angle', 'var_angle','var_width_guardCell', 
                                                        'var_length_guardCell', 'var_area_guardCell', 'wst_density','ratio_area_st_to_gc','ratio_area_to_img', 
                                                        'SEve', 'SDiv', 'SAgg'])

                        try: 
                            df.to_csv(os.path.join(str(new_path), f""+ filename_without_ext +".csv"),index=False)
                        except OSError:
                            self.closeExcel()                

                    StomGuardCell(imgPath, self.p)
                except Exception as e:
                    self.raiseError(e)
                
            self.show_info_messagebox()

        else:
            self.messagebox_define_input_path_2()
            pass           

    def Stata(self):
        """ do statistical analysis by iterate all .csv files in the output image folder,
        and generate an xlsx file containing statistical results on image level """

        self.output_folder = self.lineEdit_2.text()

        output_image_path_ = os.path.join(str(self.output_folder), "Predict_output/Output_csv")
        print(output_image_path_)
        if bool(glob.glob(output_image_path_ + "/" + '*csv')) is True:

            # for whole_stomata
            No_wst_mean = []
            No_wst_median = []
            No_wst_min = []
            No_wst_max = []

            box_w_wst_mean = []
            box_w_wst_median = []
            box_w_wst_min = []
            box_w_wst_max = []

            box_h_wst_mean = []
            box_h_wst_median = []
            box_h_wst_min = []
            box_h_wst_max = []

            area_wst_mean = []
            area_wst_median = []
            area_wst_min = []
            area_wst_max = []

            width_wst_mean = []
            width_wst_median = []
            width_wst_min = []
            width_wst_max = []

            length_wst_mean = []
            length_wst_median = []
            length_wst_min = []
            length_wst_max = []

            var_area_wst_mean = []
            var_area_wst_median = []
            var_area_wst_min = []
            var_area_wst_max = []

            var_width_wst_mean = []
            var_width_wst_median = []
            var_width_wst_min = []
            var_width_wst_max = []

            var_length_wst_mean = []
            var_length_wst_median = []
            var_length_wst_min = []
            var_length_wst_max = []    

            ## for stomata
            No_st_mean = []
            No_st_median = []
            No_st_min = []
            No_st_max = []

            box_w_st_mean = []
            box_w_st_median = []
            box_w_st_min = []
            box_w_st_max = []

            box_h_st_mean = []
            box_h_st_median = []
            box_h_st_min = []
            box_h_st_max = []

            area_st_mean = []
            area_st_median = []
            area_st_min = []
            area_st_max = []

            width_st_mean = []
            width_st_median = []
            width_st_min = []
            width_st_max = []

            length_st_mean = []
            length_st_median = []
            length_st_min = []
            length_st_max = []

            var_area_st_mean = []
            var_area_st_median = []
            var_area_st_min = []
            var_area_st_max = []

            var_width_st_mean = []
            var_width_st_median = []
            var_width_st_min = []
            var_width_st_max = []

            var_length_st_mean = []
            var_length_st_median = []
            var_length_st_min = []
            var_length_st_max = []   

            ## for guard cell
            guardCell_length_mean = []  
            guardCell_length_median = []    
            guardCell_length_min = [] 
            guardCell_length_max = []    

            guardCell_width_mean = []
            guardCell_width_median = []
            guardCell_width_min = []
            guardCell_width_max = []

            guardCell_area_mean = []
            guardCell_area_median = []
            guardCell_area_min = []
            guardCell_area_max = []

            guardCell_angle_mean = []
            guardCell_angle_median = []
            guardCell_angle_min = []
            guardCell_angle_max = []

            var_width_guardCell_mean = []
            var_width_guardCell_median = []
            var_width_guardCell_min = []
            var_width_guardCell_max = []

            var_length_guardCell_mean = []
            var_length_guardCell_median = []
            var_length_guardCell_min = []
            var_length_guardCell_max = []

            var_area_guardCell_mean = []
            var_area_guardCell_median = []
            var_area_guardCell_min = []
            var_area_guardCell_max = []

            wst_density_mean = []
            wst_density_median = []
            wst_density_min = []
            wst_density_max = []

            ratio_area_st_gc_mean = []
            ratio_area_st_gc_median = []
            ratio_area_st_gc_min = []
            ratio_area_st_gc_max = []

            ratio_area_to_img_mean = []
            ratio_area_to_img_median = []
            ratio_area_to_img_min = []
            ratio_area_to_img_max = []

            var_angle_mean = []
            var_angle_median = []
            var_angle_min = []
            var_angle_max = []

            SEve_mean = []
            SEve_median = []
            SEve_min = []
            SEve_max = []            

            SDiv_mean = []
            SDiv_median = []
            SDiv_min = []
            SDiv_max = []     

            SAgg_mean = []
            SAgg_median = []
            SAgg_min = []
            SAgg_max = []  

            Filename = []

            # Using for loop to go through all csv files
            if bool(glob.glob(output_image_path_ + '/' + 'Statistics.csv')) is True:
                os.remove(glob.glob(output_image_path_ + '/' + 'Statistics.csv')[0])
            else:
                self.img_num =1
                for file_path in glob.glob(output_image_path_ + '/' + '*.csv'):
                    
                    single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable

                    file_name = file_path[len(output_image_path_) + 1:-4]  # Extract the file name from the file path
                    Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
                    if single_csv_file.shape[0]>=4:

                        try:

                            Filename.append(",".join([str(item) for item in Split_name]))
                            # Extract the parameters from each single csv files
                            number_wst = single_csv_file["number_wst"]
                            box_w_wst = single_csv_file["box_w_wst"]
                            box_h_wst = single_csv_file["box_h_wst"]  
                            area_wst = single_csv_file["area_wst"]
                            width_wst = single_csv_file["width_wst"]
                            length_wst = single_csv_file["length_wst"]
                            var_area_wst = single_csv_file["var_area_wst"]
                            var_width_wst = single_csv_file["var_width_wst"]
                            var_length_wst = single_csv_file["var_length_wst"]
                            
                            number_st = single_csv_file["number_st"]
                            box_w_st = single_csv_file["box_w_st"]
                            box_h_st = single_csv_file["box_h_st"]  
                            area_st = single_csv_file["area_st"]
                            width_st = single_csv_file["width_st"]
                            length_st = single_csv_file["length_st"]
                            var_area_st = single_csv_file["var_area_st"]
                            var_width_st = single_csv_file["var_width_st"]
                            var_length_st = single_csv_file["var_length_st"]

                            guardCell_length = single_csv_file["guardCell_length"]
                            guardCell_width = single_csv_file["guardCell_width"]
                            guardCell_area = single_csv_file["guardCell_area"]
                            guardCell_angle = single_csv_file["guardCell_angle"]
                            var_angle = single_csv_file["var_angle"]
                            var_width_guardCell = single_csv_file["var_width_guardCell"]
                            var_length_guardCell = single_csv_file["var_length_guardCell"]
                            var_area_guardCell = single_csv_file["var_area_guardCell"]                        
                            wst_density = single_csv_file["wst_density"]
                            ratio_area_st_gc = single_csv_file["ratio_area_st_to_gc"]
                            ratio_area_to_img = single_csv_file["ratio_area_to_img"]

                            SEve = single_csv_file["SEve"]
                            SDiv = single_csv_file["SDiv"]
                            SAgg = single_csv_file["SAgg"]

                            # print(type(box_w_wst))

                            # Remove the outliers before calculating box_w_wst variance
                            Q1_box_w_wst = np.percentile(box_w_wst, 5, method = 'midpoint')
                            Q3_box_w_wst = np.percentile(box_w_wst, 95, method = 'midpoint')
                            IQR_box_w_wst = Q3_box_w_wst - Q1_box_w_wst
                            upper = np.where(box_w_wst >= (Q3_box_w_wst+1.5*IQR_box_w_wst))
                            lower = np.where(box_w_wst <= (Q1_box_w_wst-1.5*IQR_box_w_wst))
                            box_w_wst.drop(upper[0], inplace = True)
                            box_w_wst.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating box_h_wst variance
                            Q1_box_h_wst = np.percentile(box_h_wst, 2.5, method = 'midpoint')
                            Q3_box_h_wst = np.percentile(box_h_wst, 97.5,method = 'midpoint')
                            IQR_box_h_wst = Q3_box_h_wst - Q1_box_h_wst
                            upper = np.where(box_h_wst >= (Q3_box_h_wst+1.5*IQR_box_h_wst))
                            lower = np.where(box_h_wst <= (Q1_box_h_wst-1.5*IQR_box_h_wst))
                            box_h_wst.drop(upper[0], inplace = True)
                            box_h_wst.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating area_wst variance
                            Q1_area_wst = np.percentile(area_wst, 2.5, method = 'midpoint')
                            Q3_area_wst = np.percentile(area_wst, 97.5,method = 'midpoint')
                            IQR_area_wst = Q3_area_wst - Q1_area_wst
                            upper = np.where(area_wst >= (Q3_area_wst+1.5*IQR_area_wst))
                            lower = np.where(area_wst <= (Q1_area_wst-1.5*IQR_area_wst))
                            area_wst.drop(upper[0], inplace = True)
                            area_wst.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating width_wst variance
                            Q1_width_wst = np.percentile(width_wst, 2.5, method = 'midpoint')
                            Q3_width_wst = np.percentile(width_wst, 97.5,method = 'midpoint')
                            IQR_width_wst = Q3_width_wst - Q1_width_wst
                            upper = np.where(width_wst >= (Q3_width_wst+1.5*IQR_width_wst))
                            lower = np.where(width_wst <= (Q1_width_wst-1.5*IQR_width_wst))
                            width_wst.drop(upper[0], inplace = True)
                            width_wst.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating length_wst variance
                            Q1_length_wst = np.percentile(length_wst, 2.5, method = 'midpoint')
                            Q3_length_wst = np.percentile(length_wst, 97.5,method = 'midpoint')
                            IQR_length_wst = Q3_length_wst - Q1_length_wst
                            upper = np.where(length_wst >= (Q3_length_wst+1.5*IQR_length_wst))
                            lower = np.where(length_wst <= (Q1_length_wst-1.5*IQR_length_wst))
                            length_wst.drop(upper[0], inplace = True)
                            length_wst.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating var_area_wst variance
                            # Q1_var_area_wst = np.percentile(var_area_wst, 5, method = 'midpoint')
                            # Q3_var_area_wst = np.percentile(var_area_wst, 95,method = 'midpoint')
                            # IQR_var_area_wst = Q3_var_area_wst - Q1_var_area_wst
                            # upper = np.where(var_area_wst >= (Q3_var_area_wst+1.5*IQR_var_area_wst))
                            # lower = np.where(var_area_wst <= (Q1_var_area_wst-1.5*IQR_var_area_wst))
                            # var_area_wst.drop(upper[0], inplace = True)
                            # var_area_wst.drop(lower[0], inplace = True)
                            var_area_wst = var_area_wst[-1:]

                            # Remove the outliers before calculating var_width_wst variance
                            # Q1_var_width_wst = np.percentile(var_width_wst, 5, method = 'midpoint')
                            # Q3_var_width_wst = np.percentile(var_width_wst, 95,method = 'midpoint')
                            # IQR_var_width_wst = Q3_var_width_wst - Q1_var_width_wst
                            # upper = np.where(var_width_wst >= (Q3_var_width_wst+1.5*IQR_var_width_wst))
                            # lower = np.where(var_width_wst <= (Q1_var_width_wst-1.5*IQR_var_width_wst))
                            # var_width_wst.drop(upper[0], inplace = True)
                            # var_width_wst.drop(lower[0], inplace = True) 
                            var_width_wst = var_width_wst[-1:]                

                            var_angle = var_angle[-1:]   
                            # Remove the outliers before calculating var_length_wst variance
                            # Q1_var_length_wst = np.percentile(var_length_wst, 5, method = 'midpoint')
                            # Q3_var_length_wst = np.percentile(var_length_wst, 95,method = 'midpoint')
                            # IQR_var_length_wst = Q3_var_length_wst - Q1_var_length_wst
                            # upper = np.where(var_length_wst >= (Q3_var_length_wst+1.5*IQR_var_length_wst))
                            # lower = np.where(var_length_wst <= (Q1_var_length_wst-1.5*IQR_var_length_wst))
                            # var_length_wst.drop(upper[0], inplace = True)
                            # var_length_wst.drop(lower[0], inplace = True)  
                            var_length_wst = var_length_wst[-1:] 

                            # Remove the outliers before calculating box_w_st variance
                            Q1_box_w_st = np.percentile(box_w_st, 2.5, method = 'midpoint')
                            Q3_box_w_st = np.percentile(box_w_st, 97.5,method = 'midpoint')
                            IQR_box_w_st = Q3_box_w_st - Q1_box_w_st
                            upper = np.where(box_w_st >= (Q3_box_w_st+1.5*IQR_box_w_st))
                            lower = np.where(box_w_st <= (Q1_box_w_st-1.5*IQR_box_w_st))
                            box_w_st.drop(upper[0], inplace = True)
                            box_w_st.drop(lower[0], inplace = True)                               

                            # Remove the outliers before calculating box_h_st variance
                            Q1_box_h_st = np.percentile(box_h_st, 2.5, method = 'midpoint')
                            Q3_box_h_st = np.percentile(box_h_st, 97.5,method = 'midpoint')
                            IQR_box_h_st = Q3_box_h_st - Q1_box_h_st
                            upper = np.where(box_h_st >= (Q3_box_h_st+1.5*IQR_box_h_st))
                            lower = np.where(box_h_st <= (Q1_box_h_st-1.5*IQR_box_h_st))
                            box_h_st.drop(upper[0], inplace = True)
                            box_h_st.drop(lower[0], inplace = True)   

                            # Remove the outliers before calculating area_st variance
                            Q1_area_st = np.percentile(area_st, 2.5, method = 'midpoint')
                            Q3_area_st = np.percentile(area_st, 97.5,method = 'midpoint')
                            IQR_area_st = Q3_area_st - Q1_area_st
                            upper = np.where(area_st >= (Q3_area_st+1.5*IQR_area_st))
                            lower = np.where(area_st <= (Q1_area_st-1.5*IQR_area_st))
                            area_st.drop(upper[0], inplace = True)
                            area_st.drop(lower[0], inplace = True)   

                            # Remove the outliers before calculating width_st variance
                            Q1_width_st = np.percentile(width_st, 2.5, method = 'midpoint')
                            Q3_width_st = np.percentile(width_st, 97.5,method = 'midpoint')
                            IQR_width_st = Q3_width_st - Q1_width_st
                            upper = np.where(width_st >= (Q3_width_st+1.5*IQR_width_st))
                            lower = np.where(width_st <= (Q1_width_st-1.5*IQR_width_st))
                            width_st.drop(upper[0], inplace = True)
                            width_st.drop(lower[0], inplace = True)  

                            # Remove the outliers before calculating length_st variance
                            Q1_length_st = np.percentile(length_st, 2.5, method = 'midpoint')
                            Q3_length_st = np.percentile(length_st, 97.5,method = 'midpoint')
                            IQR_length_st = Q3_length_st - Q1_length_st
                            upper = np.where(length_st >= (Q3_length_st+1.5*IQR_length_st))
                            lower = np.where(length_st <= (Q1_length_st-1.5*IQR_length_st))
                            length_st.drop(upper[0], inplace = True)
                            length_st.drop(lower[0], inplace = True)

                            # Remove the outliers before calculating var_area_st variance
                            # Q1_var_area_st = np.percentile(var_area_st, 5, method = 'midpoint')
                            # Q3_var_area_st = np.percentile(var_area_st, 95,method = 'midpoint')
                            # IQR_var_area_st = Q3_var_area_st - Q1_var_area_st
                            # upper = np.where(var_area_st >= (Q3_var_area_st+1.5*IQR_var_area_st))
                            # lower = np.where(var_area_st <= (Q1_var_area_st-1.5*IQR_var_area_st))
                            # var_area_st.drop(upper[0], inplace = True)
                            # var_area_st.drop(lower[0], inplace = True)
                            var_area_st = var_area_st[-1:] 

                            # Remove the outliers before calculating var_width_st variance
                            # Q1_var_width_st = np.percentile(var_width_st, 5, method = 'midpoint')
                            # Q3_var_width_st = np.percentile(var_width_st, 95,method = 'midpoint')
                            # IQR_var_width_st = Q3_var_width_st - Q1_var_width_st
                            # upper = np.where(var_width_st >= (Q3_var_width_st+1.5*IQR_var_width_st))
                            # lower = np.where(var_width_st <= (Q1_var_width_st-1.5*IQR_var_width_st))
                            # var_width_st.drop(upper[0], inplace = True)
                            # var_width_st.drop(lower[0], inplace = True) 
                            var_width_st = var_width_st[-1:]                 

                            # Remove the outliers before calculating var_length_st variance
                            # Q1_var_length_st = np.percentile(var_length_st, 5, method = 'midpoint')
                            # Q3_var_length_st = np.percentile(var_length_st, 95,method = 'midpoint')
                            # IQR_var_length_st = Q3_var_length_st - Q1_var_length_st
                            # upper = np.where(var_length_st >= (Q3_var_length_st+1.5*IQR_var_length_st))
                            # lower = np.where(var_length_st <= (Q1_var_length_st-1.5*IQR_var_length_st))
                            # var_length_st.drop(upper[0], inplace = True)
                            # var_length_st.drop(lower[0], inplace = True) 
                            var_length_st = var_length_st[-1:] 

                            var_area_guardCell = var_area_guardCell[-1:]                                       

                            # Remove the outliers before calculating guardCell_length variance
                            Q1_guardCell_length = np.percentile(guardCell_length, 2.5, method = 'midpoint')
                            Q3_guardCell_length = np.percentile(guardCell_length, 97.5,method = 'midpoint')
                            IQR_guardCell_length = Q3_guardCell_length - Q1_guardCell_length
                            upper = np.where(guardCell_length >= (Q3_guardCell_length+1.5*IQR_guardCell_length))
                            lower = np.where(guardCell_length <= (Q1_guardCell_length-1.5*IQR_guardCell_length))
                            guardCell_length.drop(upper[0], inplace = True)
                            guardCell_length.drop(lower[0], inplace = True) 
                            
                            # Remove the outliers before calculating guardCell_width variance
                            Q1_guardCell_width = np.percentile(guardCell_width, 2.5, method = 'midpoint')
                            Q3_guardCell_width = np.percentile(guardCell_width, 97.5,method = 'midpoint')
                            IQR_guardCell_width = Q3_guardCell_width - Q1_guardCell_width
                            upper = np.where(guardCell_width >= (Q3_guardCell_width+1.5*IQR_guardCell_width))
                            lower = np.where(guardCell_width <= (Q1_guardCell_width-1.5*IQR_guardCell_width))
                            guardCell_width.drop(upper[0], inplace = True)
                            guardCell_width.drop(lower[0], inplace = True) 

                            # Remove the outliers before calculating guardCell_area variance
                            Q1_guardCell_area = np.percentile(guardCell_area, 2.5, method = 'midpoint')
                            Q3_guardCell_area = np.percentile(guardCell_area, 97.5,method = 'midpoint')
                            IQR_guardCell_area = Q3_guardCell_area - Q1_guardCell_area
                            upper = np.where(guardCell_area >= (Q3_guardCell_area+1.5*IQR_guardCell_area))
                            lower = np.where(guardCell_area <= (Q1_guardCell_area-1.5*IQR_guardCell_area))
                            guardCell_area.drop(upper[0], inplace = True)
                            guardCell_area.drop(lower[0], inplace = True) 

                            # Remove the outliers before calculating var_angle variance
                            # Q1_var_angle = np.percentile(var_angle, 5, method = 'midpoint')
                            # Q3_var_angle = np.percentile(var_angle, 95,method = 'midpoint')
                            # IQR_var_angle = Q3_var_angle - Q1_var_angle
                            # upper = np.where(var_angle >= (Q3_var_angle+1.5*IQR_var_angle))
                            # lower = np.where(var_angle <= (Q1_var_angle-1.5*IQR_var_angle))
                            # var_angle.drop(upper[0], inplace = True)
                            # var_angle.drop(lower[0], inplace = True) 

                            # Remove the outliers before calculating var_width_guardCell variance
                            # Q1_var_width_guardCell = np.percentile(var_width_guardCell, 5, method = 'midpoint')
                            # Q3_var_width_guardCell = np.percentile(var_width_guardCell, 95,method = 'midpoint')
                            # IQR_var_width_guardCell = Q3_var_width_guardCell - Q1_var_width_guardCell
                            # upper = np.where(var_width_guardCell >= (Q3_var_width_guardCell+1.5*IQR_var_width_guardCell))
                            # lower = np.where(var_width_guardCell <= (Q1_var_width_guardCell-1.5*IQR_var_width_guardCell))
                            # var_width_guardCell.drop(upper[0], inplace = True)
                            # var_width_guardCell.drop(lower[0], inplace = True)
                            var_width_guardCell = var_width_guardCell[-1:]                  

                            # Remove the outliers before calculating var_length_guardCell variance
                            # Q1_var_length_guardCell = np.percentile(var_length_guardCell, 5, method = 'midpoint')
                            # Q3_var_length_guardCell = np.percentile(var_length_guardCell, 95,method = 'midpoint')
                            # IQR_var_length_guardCell = Q3_var_length_guardCell - Q1_var_length_guardCell
                            # upper = np.where(var_length_guardCell >= (Q3_var_length_guardCell+1.5*IQR_var_length_guardCell))
                            # lower = np.where(var_length_guardCell <= (Q1_var_length_guardCell-1.5*IQR_var_length_guardCell))
                            # var_length_guardCell.drop(upper[0], inplace = True)
                            # var_length_guardCell.drop(lower[0], inplace = True) 
                            var_length_guardCell = var_length_guardCell[-1:] 

                            # Remove the outliers before calculating ratio_area_st_gc variance
                            Q1_ratio_area_st_gc = np.percentile(ratio_area_st_gc, 2.5, method = 'midpoint')
                            Q3_ratio_area_st_gc = np.percentile(ratio_area_st_gc, 97.5,method = 'midpoint')
                            IQR_ratio_area_st_gc = Q3_ratio_area_st_gc - Q1_ratio_area_st_gc
                            upper = np.where(ratio_area_st_gc >= (Q3_ratio_area_st_gc+1.5*IQR_ratio_area_st_gc))
                            lower = np.where(ratio_area_st_gc <= (Q1_ratio_area_st_gc-1.5*IQR_ratio_area_st_gc))
                            ratio_area_st_gc.drop(upper[0], inplace = True)
                            ratio_area_st_gc.drop(lower[0], inplace = True) 

                            # Remove the outliers before calculating ratio_area_to_img variance
                            ratio_area_to_img = ratio_area_to_img[-1:] 

                            # calculate the median, mean, variance of each leaf stomata number, stomata area

                            No_wst_mean_ = np.mean(number_wst)
                            No_wst_median_ = np.median(number_wst)
                            No_wst_min_ = min(number_wst)
                            No_wst_max_ = max(number_wst)

                            box_w_wst_mean_ = np.mean(box_w_wst)
                            box_w_wst_median_ = np.median(box_w_wst)
                            box_w_wst_min_ = min(box_w_wst)
                            box_w_wst_max_ = max(box_w_wst)

                            box_h_wst_mean_ = np.mean(box_h_wst)
                            box_h_wst_median_ = np.median(box_h_wst)
                            box_h_wst_min_ = min(box_h_wst)
                            box_h_wst_max_ = max(box_h_wst)

                            area_wst_mean_ = np.mean(area_wst)
                            area_wst_median_ = np.median(area_wst)
                            area_wst_min_ = min(area_wst)
                            area_wst_max_ = max(area_wst)

                            width_wst_mean_ = np.mean(width_wst)
                            width_wst_median_ = np.median(width_wst)
                            width_wst_min_ = min(width_wst)
                            width_wst_max_ = max(width_wst)

                            length_wst_mean_ = np.mean(length_wst)
                            length_wst_median_ = np.median(length_wst)
                            length_wst_min_ = min(length_wst)
                            length_wst_max_ = max(length_wst)

                            var_area_wst_mean_ = np.mean(var_area_wst)
                            var_area_wst_median_ = np.median(var_area_wst)
                            var_area_wst_min_ = min(var_area_wst)
                            var_area_wst_max_ = max(var_area_wst)

                            var_width_wst_mean_ = np.mean(var_width_wst)
                            var_width_wst_median_ = np.median(var_width_wst)
                            var_width_wst_min_ = min(var_width_wst)
                            var_width_wst_max_ = max(var_width_wst)

                            var_length_wst_mean_ = np.mean(var_length_wst)
                            var_length_wst_median_ = np.median(var_length_wst)
                            var_length_wst_min_ = min(var_length_wst)
                            var_length_wst_max_ = max(var_length_wst)  

                            ## for stomata
                            No_st_mean_ = np.mean(number_st)
                            No_st_median_ = np.median(number_st)
                            No_st_min_ = min(number_st)
                            No_st_max_ = max(number_st)

                            box_w_st_mean_ = np.mean(box_w_st)
                            box_w_st_median_ = np.median(box_w_st)
                            box_w_st_min_ = min(box_w_st)
                            box_w_st_max_ = max(box_w_st)

                            box_h_st_mean_ = np.mean(box_h_st)
                            box_h_st_median_ = np.median(box_h_st)
                            box_h_st_min_ = min(box_h_st)
                            box_h_st_max_ = max(box_h_st)

                            area_st_mean_ = np.mean(area_st)
                            area_st_median_ = np.median(area_st)
                            area_st_min_ = min(area_st)
                            area_st_max_ = max(area_st)

                            width_st_mean_ = np.mean(width_st)
                            width_st_median_ = np.median(width_st)
                            width_st_min_ = min(width_st)
                            width_st_max_ = max(width_st)

                            length_st_mean_ = np.mean(length_st)
                            length_st_median_ = np.median(length_st)
                            length_st_min_ = min(length_st)
                            length_st_max_ = max(length_st)

                            var_area_st_mean_ = np.mean(var_area_st)
                            var_area_st_median_ = np.median(var_area_st)
                            var_area_st_min_ = min(var_area_st)
                            var_area_st_max_ = max(var_area_st)

                            var_width_st_mean_ = np.mean(var_width_st)
                            var_width_st_median_ = np.median(var_width_st)
                            var_width_st_min_ = min(var_width_st)
                            var_width_st_max_ = max(var_width_st)

                            var_length_st_mean_ = np.mean(var_length_st)
                            var_length_st_median_ = np.median(var_length_st)
                            var_length_st_min_ = min(var_length_st)
                            var_length_st_max_ = max(var_length_st)  

                            ## for guard cell
                            guardCell_length_mean_ = np.mean(guardCell_length)  
                            guardCell_length_median_ = np.median(guardCell_length)  
                            guardCell_length_min_ = min(guardCell_length)
                            guardCell_length_max_ = max(guardCell_length)    

                            guardCell_width_mean_ = np.mean(guardCell_width)
                            guardCell_width_median_ = np.median(guardCell_width)
                            guardCell_width_min_ = min(guardCell_width)
                            guardCell_width_max_ = max(guardCell_width)

                            guardCell_area_mean_ = np.mean(guardCell_area)
                            guardCell_area_median_ = np.median(guardCell_area)
                            guardCell_area_min_ = min(guardCell_area)
                            guardCell_area_max_ = max(guardCell_area)

                            guardCell_angle_mean_ = np.mean(guardCell_angle)
                            guardCell_angle_median_ = np.median(guardCell_angle)
                            guardCell_angle_min_ = min(guardCell_angle)
                            guardCell_angle_max_ = max(guardCell_angle)

                            var_width_guardCell_mean_ = np.mean(var_width_guardCell)
                            var_width_guardCell_median_ = np.median(var_width_guardCell)
                            var_width_guardCell_min_ = min(var_width_guardCell)
                            var_width_guardCell_max_ = max(var_width_guardCell)

                            var_length_guardCell_mean_ = np.mean(var_length_guardCell)
                            var_length_guardCell_median_ = np.median(var_length_guardCell)
                            var_length_guardCell_min_ = min(var_length_guardCell)
                            var_length_guardCell_max_ = max(var_length_guardCell)

                            var_area_guardCell_mean_ = np.mean(var_area_guardCell)
                            var_area_guardCell_median_ = np.median(var_area_guardCell)
                            var_area_guardCell_min_ = min(var_area_guardCell)
                            var_area_guardCell_max_ = max(var_area_guardCell)                          

                            wst_density_mean_ = np.mean(wst_density)
                            wst_density_median_ = np.median(wst_density)
                            wst_density_min_ = min(wst_density)
                            wst_density_max_ = max(wst_density)

                            ratio_area_st_gc_mean_ = np.mean(ratio_area_st_gc)
                            ratio_area_st_gc_median_ = np.median(ratio_area_st_gc)
                            ratio_area_st_gc_min_ = min(ratio_area_st_gc)
                            ratio_area_st_gc_max_ = max(ratio_area_st_gc)

                            ratio_area_to_img_mean_ = np.mean(ratio_area_to_img)
                            ratio_area_to_img_median_ = np.median(ratio_area_to_img)
                            ratio_area_to_img_min_ = min(ratio_area_to_img)
                            ratio_area_to_img_max_ = max(ratio_area_to_img)

                            var_angle_mean_ = np.mean(var_angle)
                            var_angle_median_ = np.median(var_angle)
                            var_angle_min_ = min(var_angle)
                            var_angle_max_ = max(var_angle)

                            SEve_mean_ = np.mean(SEve)
                            SEve_median_ = np.median(SEve)
                            SEve_min_ = min(SEve)
                            SEve_max_ = max(SEve)

                            SDiv_mean_ = np.mean(SDiv)
                            SDiv_median_ = np.median(SDiv)
                            SDiv_min_ = min(SDiv)
                            SDiv_max_ = max(SDiv)

                            SAgg_mean_ = np.mean(SAgg)
                            SAgg_median_ = np.median(SAgg)
                            SAgg_min_ = min(SAgg)
                            SAgg_max_ = max(SAgg)

                            ## append these measurements to the lists
                            # for whole_stomata
                            No_wst_mean.append(No_wst_mean_)
                            No_wst_median.append(No_wst_median_)
                            No_wst_min.append(No_wst_min_)
                            No_wst_max.append(No_wst_max_)

                            box_w_wst_mean.append(box_w_wst_mean_)
                            box_w_wst_median.append(box_w_wst_median_)
                            box_w_wst_min.append(box_w_wst_min_)
                            box_w_wst_max.append(box_w_wst_max_)

                            box_h_wst_mean.append(box_h_wst_mean_)
                            box_h_wst_median.append(box_h_wst_median_)
                            box_h_wst_min.append(box_h_wst_min_)
                            box_h_wst_max.append(box_h_wst_max_)

                            area_wst_mean.append(area_wst_mean_)
                            area_wst_median.append(area_wst_median_)
                            area_wst_min.append(area_wst_min_)
                            area_wst_max.append(area_wst_max_)

                            width_wst_mean.append(width_wst_mean_)
                            width_wst_median.append(width_wst_median_)
                            width_wst_min.append(width_wst_min_)
                            width_wst_max.append(width_wst_max_)

                            length_wst_mean.append(length_wst_mean_)
                            length_wst_median.append(length_wst_median_)
                            length_wst_min.append(length_wst_min_)
                            length_wst_max.append(length_wst_max_)

                            var_area_wst_mean.append(var_area_wst_mean_)
                            var_area_wst_median.append(var_area_wst_median_)
                            var_area_wst_min.append(var_area_wst_min_)
                            var_area_wst_max.append(var_area_wst_max_)

                            var_width_wst_mean.append(var_width_wst_mean_)
                            var_width_wst_median.append(var_width_wst_median_)
                            var_width_wst_min.append(var_width_wst_min_)
                            var_width_wst_max.append(var_width_wst_max_)

                            var_length_wst_mean.append(var_length_wst_mean_)
                            var_length_wst_median.append(var_length_wst_median_)
                            var_length_wst_min.append(var_length_wst_min_)
                            var_length_wst_max.append(var_length_wst_max_)    

                            ## for stomata
                            No_st_mean.append(No_st_mean_)
                            No_st_median.append(No_st_median_)
                            No_st_min.append(No_st_min_)
                            No_st_max.append(No_st_max_)

                            box_w_st_mean.append(box_w_st_mean_)
                            box_w_st_median.append(box_w_st_median_)
                            box_w_st_min.append(box_w_st_min_)
                            box_w_st_max.append(box_w_st_max_)

                            box_h_st_mean.append(box_h_st_mean_)
                            box_h_st_median.append(box_h_st_median_)
                            box_h_st_min.append(box_h_st_min_)
                            box_h_st_max.append(box_h_st_max_)

                            area_st_mean.append(area_st_mean_)
                            area_st_median.append(area_st_median_)
                            area_st_min.append(area_st_min_)
                            area_st_max.append(area_st_max_)

                            width_st_mean.append(width_st_mean_)
                            width_st_median.append(width_st_median_)
                            width_st_min.append(width_st_min_)
                            width_st_max.append(width_st_max_)

                            length_st_mean.append(length_st_mean_)
                            length_st_median.append(length_st_median_)
                            length_st_min.append(length_st_min_)
                            length_st_max.append(length_st_max_)

                            var_area_st_mean.append(var_area_st_mean_)
                            var_area_st_median.append(var_area_st_median_)
                            var_area_st_min.append(var_area_st_min_)
                            var_area_st_max.append(var_area_st_max_)

                            var_width_st_mean.append(var_width_st_mean_)
                            var_width_st_median.append(var_width_st_median_)
                            var_width_st_min.append(var_width_st_min_)
                            var_width_st_max.append(var_width_st_max_)

                            var_length_st_mean.append(var_length_st_mean_)
                            var_length_st_median.append(var_length_st_median_)
                            var_length_st_min.append(var_length_st_min_)
                            var_length_st_max.append(var_length_st_max_)  

                            ## for guard cell
                            guardCell_length_mean.append(guardCell_length_mean_) 
                            guardCell_length_median.append(guardCell_length_median_)  
                            guardCell_length_min.append(guardCell_length_min_)
                            guardCell_length_max.append(guardCell_length_max_)    

                            guardCell_width_mean.append(guardCell_width_mean_)
                            guardCell_width_median.append(guardCell_width_median_)
                            guardCell_width_min.append(guardCell_width_min_)
                            guardCell_width_max.append(guardCell_width_max_)

                            guardCell_area_mean.append(guardCell_area_mean_)
                            guardCell_area_median.append(guardCell_area_median_)
                            guardCell_area_min.append(guardCell_area_min_)
                            guardCell_area_max.append(guardCell_area_max_)

                            guardCell_angle_mean.append(guardCell_angle_mean_)
                            guardCell_angle_median.append(guardCell_angle_median_)
                            guardCell_angle_min.append(guardCell_angle_min_)
                            guardCell_angle_max.append(guardCell_angle_max_)

                            var_width_guardCell_mean.append(var_width_guardCell_mean_)
                            var_width_guardCell_median.append(var_width_guardCell_median_)
                            var_width_guardCell_min.append(var_width_guardCell_min_)
                            var_width_guardCell_max.append(var_width_guardCell_max_)

                            var_length_guardCell_mean.append(var_length_guardCell_mean_)
                            var_length_guardCell_median.append(var_length_guardCell_median_)
                            var_length_guardCell_min.append(var_length_guardCell_min_)
                            var_length_guardCell_max.append(var_length_guardCell_max_)

                            var_area_guardCell_mean.append(var_area_guardCell_mean_)
                            var_area_guardCell_median.append(var_area_guardCell_median_)
                            var_area_guardCell_min.append(var_area_guardCell_min_)
                            var_area_guardCell_max.append(var_area_guardCell_max_)                        

                            wst_density_mean.append(wst_density_mean_)
                            wst_density_median.append(wst_density_median_)
                            wst_density_min.append(wst_density_min_)
                            wst_density_max.append(wst_density_max_)

                            ratio_area_st_gc_mean.append(ratio_area_st_gc_mean_)
                            ratio_area_st_gc_median.append(ratio_area_st_gc_median_)
                            ratio_area_st_gc_min.append(ratio_area_st_gc_min_)
                            ratio_area_st_gc_max.append(ratio_area_st_gc_max_)

                            var_angle_mean.append(var_angle_mean_)
                            var_angle_median.append(var_angle_median_)
                            var_angle_min.append(var_angle_min_)
                            var_angle_max.append(var_angle_max_)

                            ratio_area_to_img_mean.append(ratio_area_to_img_mean_)
                            ratio_area_to_img_median.append(ratio_area_to_img_median_)
                            ratio_area_to_img_min.append(ratio_area_to_img_min_)
                            ratio_area_to_img_max.append(ratio_area_to_img_max_)

                            SEve_mean.append(SEve_mean_)
                            SEve_median.append(SEve_median_)
                            SEve_min.append(SEve_min_)
                            SEve_max.append(SEve_max_)

                            SDiv_mean.append(SDiv_mean_)
                            SDiv_median.append(SDiv_median_)
                            SDiv_min.append(SDiv_min_)
                            SDiv_max.append(SDiv_max_)

                            SAgg_mean.append(SAgg_mean_)
                            SAgg_median.append(SAgg_median_)
                            SAgg_min.append(SAgg_min_)
                            SAgg_max.append(SAgg_max_)
                            
                            self.progressBar.setValue(int((self.img_num / len(glob.glob(output_image_path_ + '/' + '*.csv'))*100)))
                            QtWidgets.QApplication.processEvents()

                            self.img_num +=1

                        except KeyError:
                            self.keyError()
                            Filename.pop()
                            self.img_num +=2

                    else:
                        self.img_num +=1
                        pass
        
                # Convert all lists into pd.series
                Filename = pd.Series(Filename, dtype=pd.StringDtype(), name="Filename")
                
                No_wst_mean = pd.Series(No_wst_mean, dtype=pd.Float64Dtype(), name="No_wst_mean").map('{:,.0f}'.format)
                No_wst_median = pd.Series(No_wst_median, dtype=pd.Float64Dtype(), name="No_wst_median").map('{:,.0f}'.format)
                No_wst_min = pd.Series(No_wst_min, dtype=pd.Float64Dtype(), name="No_wst_min").map('{:,.0f}'.format)
                No_wst_max = pd.Series(No_wst_max, dtype=pd.Float64Dtype(), name="No_wst_max").map('{:,.0f}'.format)

                box_w_wst_mean = pd.Series(box_w_wst_mean, dtype=pd.Float64Dtype(), name="box_w_wst_mean").map('{:,.0f}'.format)
                box_w_wst_median = pd.Series(box_w_wst_median, dtype=pd.Float64Dtype(), name="box_w_wst_median").map('{:,.0f}'.format)
                box_w_wst_min = pd.Series(box_w_wst_min, dtype=pd.Float64Dtype(), name="box_w_wst_min").map('{:,.0f}'.format)
                box_w_wst_max = pd.Series(box_w_wst_max, dtype=pd.Float64Dtype(), name="box_w_wst_max").map('{:,.0f}'.format)

                box_h_wst_mean = pd.Series(box_h_wst_mean, dtype=pd.Float64Dtype(), name="box_h_wst_mean").map('{:,.0f}'.format)
                box_h_wst_median = pd.Series(box_h_wst_median, dtype=pd.Float64Dtype(), name="box_h_wst_median").map('{:,.0f}'.format)
                box_h_wst_min = pd.Series(box_h_wst_min, dtype=pd.Float64Dtype(), name="box_h_wst_min").map('{:,.0f}'.format)
                box_h_wst_max = pd.Series(box_h_wst_max, dtype=pd.Float64Dtype(), name="box_h_wst_max").map('{:,.0f}'.format)

                area_wst_mean = pd.Series(area_wst_mean, dtype=pd.Float64Dtype(), name="area_wst_mean").map('{:,.0f}'.format)
                area_wst_median = pd.Series(area_wst_median, dtype=pd.Float64Dtype(), name="area_wst_median").map('{:,.0f}'.format)
                area_wst_min = pd.Series(area_wst_min, dtype=pd.Float64Dtype(), name="area_wst_min").map('{:,.0f}'.format)
                area_wst_max = pd.Series(area_wst_max, dtype=pd.Float64Dtype(), name="area_wst_max").map('{:,.0f}'.format)

                width_wst_mean = pd.Series(width_wst_mean, dtype=pd.Float64Dtype(), name="width_wst_mean").map('{:,.0f}'.format)
                width_wst_median = pd.Series(width_wst_median, dtype=pd.Float64Dtype(), name="width_wst_median").map('{:,.0f}'.format)
                width_wst_min = pd.Series(width_wst_min, dtype=pd.Float64Dtype(), name="width_wst_min").map('{:,.0f}'.format)
                width_wst_max = pd.Series(width_wst_max, dtype=pd.Float64Dtype(), name="width_wst_max").map('{:,.0f}'.format)

                length_wst_mean = pd.Series(length_wst_mean, dtype=pd.Float64Dtype(), name="length_wst_mean").map('{:,.0f}'.format)
                length_wst_median = pd.Series(length_wst_median, dtype=pd.Float64Dtype(), name="length_wst_median").map('{:,.0f}'.format)
                length_wst_min = pd.Series(length_wst_min, dtype=pd.Float64Dtype(), name="length_wst_min").map('{:,.0f}'.format)
                length_wst_max = pd.Series(length_wst_max, dtype=pd.Float64Dtype(), name="length_wst_max").map('{:,.0f}'.format)

                var_area_wst_mean = pd.Series(var_area_wst_mean, dtype=pd.Float64Dtype(), name="var_area_wst_mean").map('{:,.0f}'.format)
                var_area_wst_median = pd.Series(var_area_wst_median, dtype=pd.Float64Dtype(), name="var_area_wst_median").map('{:,.0f}'.format)
                var_area_wst_min = pd.Series(var_area_wst_min, dtype=pd.Float64Dtype(), name="var_area_wst_min").map('{:,.0f}'.format)
                var_area_wst_max = pd.Series(var_area_wst_max, dtype=pd.Float64Dtype(), name="var_area_wst_max").map('{:,.0f}'.format)

                var_width_wst_mean = pd.Series(var_width_wst_mean, dtype=pd.Float64Dtype(), name="var_width_wst_mean").map('{:,.0f}'.format)
                var_width_wst_median = pd.Series(var_width_wst_median, dtype=pd.Float64Dtype(), name="var_width_wst_median").map('{:,.0f}'.format)
                var_width_wst_min = pd.Series(var_width_wst_min, dtype=pd.Float64Dtype(), name="var_width_wst_min").map('{:,.0f}'.format)
                var_width_wst_max = pd.Series(var_width_wst_max, dtype=pd.Float64Dtype(), name="var_width_wst_max").map('{:,.0f}'.format)       

                var_length_wst_mean = pd.Series(var_length_wst_mean, dtype=pd.Float64Dtype(), name="var_length_wst_mean").map('{:,.0f}'.format)
                var_length_wst_median = pd.Series(var_length_wst_median, dtype=pd.Float64Dtype(), name="var_length_wst_median").map('{:,.0f}'.format)
                var_length_wst_min = pd.Series(var_length_wst_min, dtype=pd.Float64Dtype(), name="var_length_wst_min").map('{:,.0f}'.format)
                var_length_wst_max = pd.Series(var_length_wst_max, dtype=pd.Float64Dtype(), name="var_length_wst_max").map('{:,.0f}'.format)   

                No_st_mean = pd.Series(No_st_mean, dtype=pd.Float64Dtype(), name="No_st_mean").map('{:,.0f}'.format)
                No_st_median = pd.Series(No_st_median, dtype=pd.Float64Dtype(), name="No_st_median").map('{:,.0f}'.format)
                No_st_min = pd.Series(No_st_min, dtype=pd.Float64Dtype(), name="No_st_min").map('{:,.0f}'.format)
                No_st_max = pd.Series(No_st_max, dtype=pd.Float64Dtype(), name="No_st_max").map('{:,.0f}'.format)   

                box_w_st_mean = pd.Series(box_w_st_mean, dtype=pd.Float64Dtype(), name="box_w_st_mean").map('{:,.0f}'.format)
                box_w_st_median = pd.Series(box_w_st_median, dtype=pd.Float64Dtype(), name="box_w_st_median").map('{:,.0f}'.format)
                box_w_st_min = pd.Series(box_w_st_min, dtype=pd.Float64Dtype(), name="box_w_st_min").map('{:,.0f}'.format)
                box_w_st_max = pd.Series(box_w_st_max, dtype=pd.Float64Dtype(), name="box_w_st_max").map('{:,.0f}'.format)   

                box_h_st_mean = pd.Series(box_h_st_mean, dtype=pd.Float64Dtype(), name="box_h_st_mean").map('{:,.0f}'.format)
                box_h_st_median = pd.Series(box_h_st_median, dtype=pd.Float64Dtype(), name="box_h_st_median").map('{:,.0f}'.format)
                box_h_st_min = pd.Series(box_h_st_min, dtype=pd.Float64Dtype(), name="box_h_st_min").map('{:,.0f}'.format)
                box_h_st_max = pd.Series(box_h_st_max, dtype=pd.Float64Dtype(), name="box_h_st_max").map('{:,.0f}'.format)   

                area_st_mean = pd.Series(area_st_mean, dtype=pd.Float64Dtype(), name="area_st_mean").map('{:,.0f}'.format)
                area_st_median = pd.Series(area_st_median, dtype=pd.Float64Dtype(), name="area_st_median").map('{:,.0f}'.format)
                area_st_min = pd.Series(area_st_min, dtype=pd.Float64Dtype(), name="area_st_min").map('{:,.0f}'.format)
                area_st_max = pd.Series(area_st_max, dtype=pd.Float64Dtype(), name="area_st_max").map('{:,.0f}'.format)   

                width_st_mean = pd.Series(width_st_mean, dtype=pd.Float64Dtype(), name="width_st_mean").map('{:,.0f}'.format)
                width_st_median = pd.Series(width_st_median, dtype=pd.Float64Dtype(), name="width_st_median").map('{:,.0f}'.format)
                width_st_min = pd.Series(width_st_min, dtype=pd.Float64Dtype(), name="width_st_min").map('{:,.0f}'.format)
                width_st_max = pd.Series(width_st_max, dtype=pd.Float64Dtype(), name="width_st_max").map('{:,.0f}'.format)  

                length_st_mean = pd.Series(length_st_mean, dtype=pd.Float64Dtype(), name="length_st_mean").map('{:,.0f}'.format)
                length_st_median = pd.Series(length_st_median, dtype=pd.Float64Dtype(), name="length_st_median").map('{:,.0f}'.format)
                length_st_min = pd.Series(length_st_min, dtype=pd.Float64Dtype(), name="length_st_min").map('{:,.0f}'.format)
                length_st_max = pd.Series(length_st_max, dtype=pd.Float64Dtype(), name="length_st_max").map('{:,.0f}'.format)   

                var_area_st_mean = pd.Series(var_area_st_mean, dtype=pd.Float64Dtype(), name="var_area_st_mean").map('{:,.0f}'.format)
                var_area_st_median = pd.Series(var_area_st_median, dtype=pd.Float64Dtype(), name="var_area_st_median").map('{:,.0f}'.format)
                var_area_st_min = pd.Series(var_area_st_min, dtype=pd.Float64Dtype(), name="var_area_st_min").map('{:,.0f}'.format)
                var_area_st_max = pd.Series(var_area_st_max, dtype=pd.Float64Dtype(), name="var_area_st_max").map('{:,.0f}'.format)   

                var_width_st_mean = pd.Series(var_width_st_mean, dtype=pd.Float64Dtype(), name="var_width_st_mean").map('{:,.0f}'.format)
                var_width_st_median = pd.Series(var_width_st_median, dtype=pd.Float64Dtype(), name="var_width_st_median").map('{:,.0f}'.format)
                var_width_st_min = pd.Series(var_width_st_min, dtype=pd.Float64Dtype(), name="var_width_st_min").map('{:,.0f}'.format)
                var_width_st_max = pd.Series(var_width_st_max, dtype=pd.Float64Dtype(), name="var_width_st_max").map('{:,.0f}'.format)   

                var_length_st_mean = pd.Series(var_length_st_mean, dtype=pd.Float64Dtype(), name="var_length_st_mean").map('{:,.0f}'.format)
                var_length_st_median = pd.Series(var_length_st_median, dtype=pd.Float64Dtype(), name="var_length_st_median").map('{:,.0f}'.format)
                var_length_st_min = pd.Series(var_length_st_min, dtype=pd.Float64Dtype(), name="var_length_st_min").map('{:,.0f}'.format)
                var_length_st_max = pd.Series(var_length_st_max, dtype=pd.Float64Dtype(), name="var_length_st_max").map('{:,.0f}'.format)   

                guardCell_length_mean = pd.Series(guardCell_length_mean, dtype=pd.Float64Dtype(), name="guardCell_length_mean").map('{:,.0f}'.format)
                guardCell_length_median = pd.Series(guardCell_length_median, dtype=pd.Float64Dtype(), name="guardCell_length_median").map('{:,.0f}'.format)
                guardCell_length_min = pd.Series(guardCell_length_min, dtype=pd.Float64Dtype(), name="guardCell_length_min").map('{:,.0f}'.format)
                guardCell_length_max = pd.Series(guardCell_length_max, dtype=pd.Float64Dtype(), name="guardCell_length_max").map('{:,.0f}'.format)   

                guardCell_width_mean = pd.Series(guardCell_width_mean, dtype=pd.Float64Dtype(), name="guardCell_width_mean").map('{:,.0f}'.format)
                guardCell_width_median = pd.Series(guardCell_width_median, dtype=pd.Float64Dtype(), name="guardCell_width_median").map('{:,.0f}'.format)
                guardCell_width_min = pd.Series(guardCell_width_min, dtype=pd.Float64Dtype(), name="guardCell_width_min").map('{:,.0f}'.format)
                guardCell_width_max = pd.Series(guardCell_width_max, dtype=pd.Float64Dtype(), name="guardCell_width_max").map('{:,.0f}'.format)  

                guardCell_area_mean = pd.Series(guardCell_area_mean, dtype=pd.Float64Dtype(), name="guardCell_area_mean").map('{:,.0f}'.format)
                guardCell_area_median = pd.Series(guardCell_area_median, dtype=pd.Float64Dtype(), name="guardCell_area_median").map('{:,.0f}'.format)
                guardCell_area_min = pd.Series(guardCell_area_min, dtype=pd.Float64Dtype(), name="guardCell_area_min").map('{:,.0f}'.format)
                guardCell_area_max = pd.Series(guardCell_area_max, dtype=pd.Float64Dtype(), name="guardCell_area_max").map('{:,.0f}'.format) 

                guardCell_angle_mean = pd.Series(guardCell_angle_mean, dtype=pd.Float64Dtype(), name="guardCell_angle_mean").map('{:,.0f}'.format)
                guardCell_angle_median = pd.Series(guardCell_angle_median, dtype=pd.Float64Dtype(), name="guardCell_angle_median").map('{:,.0f}'.format)
                guardCell_angle_min = pd.Series(guardCell_angle_min, dtype=pd.Float64Dtype(), name="guardCell_angle_min").map('{:,.0f}'.format)
                guardCell_angle_max = pd.Series(guardCell_angle_max, dtype=pd.Float64Dtype(), name="guardCell_angle_max").map('{:,.0f}'.format)  

                var_width_guardCell_mean = pd.Series(var_width_guardCell_mean, dtype=pd.Float64Dtype(), name="var_width_guardCell_mean").map('{:,.0f}'.format)
                var_width_guardCell_median = pd.Series(var_width_guardCell_median, dtype=pd.Float64Dtype(), name="var_width_guardCell_median").map('{:,.0f}'.format)
                var_width_guardCell_min = pd.Series(var_width_guardCell_min, dtype=pd.Float64Dtype(), name="var_width_guardCell_min").map('{:,.0f}'.format)
                var_width_guardCell_max = pd.Series(var_width_guardCell_max, dtype=pd.Float64Dtype(), name="var_width_guardCell_max").map('{:,.0f}'.format)  

                var_length_guardCell_mean = pd.Series(var_length_guardCell_mean, dtype=pd.Float64Dtype(), name="var_length_guardCell_mean").map('{:,.0f}'.format)
                var_length_guardCell_median = pd.Series(var_length_guardCell_median, dtype=pd.Float64Dtype(), name="var_length_guardCell_median").map('{:,.0f}'.format)
                var_length_guardCell_min = pd.Series(var_length_guardCell_min, dtype=pd.Float64Dtype(), name="var_length_guardCell_min").map('{:,.0f}'.format)
                var_length_guardCell_max = pd.Series(var_length_guardCell_max, dtype=pd.Float64Dtype(), name="var_length_guardCell_max").map('{:,.0f}'.format)  

                var_area_guardCell_mean = pd.Series(var_area_guardCell_mean, dtype=pd.Float64Dtype(), name="var_area_guardCell_mean").map('{:,.0f}'.format)
                var_area_guardCell_median = pd.Series(var_area_guardCell_median, dtype=pd.Float64Dtype(), name="var_area_guardCell_median").map('{:,.0f}'.format)
                var_area_guardCell_min = pd.Series(var_area_guardCell_min, dtype=pd.Float64Dtype(), name="var_area_guardCell_min").map('{:,.0f}'.format)
                var_area_guardCell_max = pd.Series(var_area_guardCell_max, dtype=pd.Float64Dtype(), name="var_area_guardCell_max").map('{:,.0f}'.format)                 

                wst_density_mean = pd.Series(wst_density_mean, dtype=pd.Float64Dtype(), name="wst_density_mean").map('{:,.0f}'.format)
                wst_density_median = pd.Series(wst_density_median, dtype=pd.Float64Dtype(), name="wst_density_median").map('{:,.0f}'.format)
                wst_density_min = pd.Series(wst_density_min, dtype=pd.Float64Dtype(), name="wst_density_min").map('{:,.0f}'.format)
                wst_density_max = pd.Series(wst_density_max, dtype=pd.Float64Dtype(), name="wst_density_max").map('{:,.0f}'.format)  

                ratio_area_st_gc_mean = pd.Series(ratio_area_st_gc_mean, dtype=pd.Float64Dtype(), name="ratio_area_st_to_gc_mean").map('{:,.3f}'.format)
                ratio_area_st_gc_median = pd.Series(ratio_area_st_gc_median, dtype=pd.Float64Dtype(), name="ratio_area_st_to_gc_median").map('{:,.3f}'.format)
                ratio_area_st_gc_min = pd.Series(ratio_area_st_gc_min, dtype=pd.Float64Dtype(), name="ratio_area_st_to_gc_min").map('{:,.3f}'.format)
                ratio_area_st_gc_max = pd.Series(ratio_area_st_gc_max, dtype=pd.Float64Dtype(), name="ratio_area_st_to_gc_max").map('{:,.3f}'.format) 

                ratio_area_to_img_mean = pd.Series(ratio_area_to_img_mean, dtype=pd.Float64Dtype(), name="ratio_area_to_img_mean").map('{:,.3f}'.format)
                ratio_area_to_img_median = pd.Series(ratio_area_to_img_median, dtype=pd.Float64Dtype(), name="ratio_area_to_img_median").map('{:,.3f}'.format)
                ratio_area_to_img_min = pd.Series(ratio_area_to_img_min, dtype=pd.Float64Dtype(), name="ratio_area_to_img_min").map('{:,.3f}'.format)
                ratio_area_to_img_max = pd.Series(ratio_area_to_img_max, dtype=pd.Float64Dtype(), name="ratio_area_to_img_max").map('{:,.3f}'.format) 

                var_angle_mean = pd.Series(var_angle_mean, dtype=pd.Float64Dtype(), name="var_angle_mean").map('{:,.0f}'.format)
                var_angle_median = pd.Series(var_angle_median, dtype=pd.Float64Dtype(), name="var_angle_median").map('{:,.0f}'.format)
                var_angle_min = pd.Series(var_angle_min, dtype=pd.Float64Dtype(), name="var_angle_min").map('{:,.0f}'.format)
                var_angle_max = pd.Series(var_angle_max, dtype=pd.Float64Dtype(), name="var_angle_max").map('{:,.0f}'.format)  

                SEve_mean = pd.Series(SEve_mean, dtype=pd.Float64Dtype(), name="SEve_mean").map('{:,.4f}'.format)
                SEve_median = pd.Series(SEve_median, dtype=pd.Float64Dtype(), name="SEve_median").map('{:,.4f}'.format)
                SEve_min = pd.Series(SEve_min, dtype=pd.Float64Dtype(), name="SEve_min").map('{:,.4f}'.format)
                SEve_max = pd.Series(SEve_max, dtype=pd.Float64Dtype(), name="SEve_max").map('{:,.4f}'.format) 

                SDiv_mean = pd.Series(SDiv_mean, dtype=pd.Float64Dtype(), name="SDiv_mean").map('{:,.4f}'.format)
                SDiv_median = pd.Series(SDiv_median, dtype=pd.Float64Dtype(), name="SDiv_median").map('{:,.4f}'.format)
                SDiv_min = pd.Series(SDiv_min, dtype=pd.Float64Dtype(), name="SDiv_min").map('{:,.4f}'.format)
                SDiv_max = pd.Series(SDiv_max, dtype=pd.Float64Dtype(), name="SDiv_max").map('{:,.4f}'.format) 

                SAgg_mean = pd.Series(SAgg_mean, dtype=pd.Float64Dtype(), name="SAgg_mean").map('{:,.4f}'.format)
                SAgg_median = pd.Series(SAgg_median, dtype=pd.Float64Dtype(), name="SAgg_median").map('{:,.4f}'.format)
                SAgg_min = pd.Series(SAgg_min, dtype=pd.Float64Dtype(), name="SAgg_min").map('{:,.4f}'.format)
                SAgg_max = pd.Series(SAgg_max, dtype=pd.Float64Dtype(), name="SAgg_max").map('{:,.4f}'.format) 

                # Put all extracted parameters into a data frame
                Output_data = pd.concat(
                    [Filename, 
                    No_wst_mean, 
                    No_wst_median,
                    No_wst_min, 
                    No_wst_max, 
                    box_w_wst_mean, 
                    box_w_wst_median,
                    box_w_wst_min, 
                    box_w_wst_max, 
                    box_h_wst_mean, 
                    box_h_wst_median, 
                    box_h_wst_min,
                    box_h_wst_max,
                    area_wst_mean,
                    area_wst_median,
                    area_wst_min,
                    area_wst_max,
                    width_wst_mean,
                    width_wst_median,
                    width_wst_min,
                    width_wst_max,
                    length_wst_mean,
                    length_wst_median,
                    length_wst_min,
                    length_wst_max,
                    var_area_wst_mean,
                    var_area_wst_median,
                    var_area_wst_min,
                    var_area_wst_max,
                    var_width_wst_mean,
                    var_width_wst_median,
                    var_width_wst_min,
                    var_width_wst_max,
                    var_length_wst_mean,
                    var_length_wst_median,
                    var_length_wst_min,
                    var_length_wst_max,
                    No_st_mean,
                    No_st_median,
                    No_st_min,
                    No_st_max,
                    box_w_st_mean,
                    box_w_st_median,
                    box_w_st_min,
                    box_w_st_max,
                    box_h_st_mean,
                    box_h_st_median,
                    box_h_st_min,
                    box_h_st_max,
                    area_st_mean,
                    area_st_median,
                    area_st_min,
                    area_st_max,
                    width_st_mean,
                    width_st_median,
                    width_st_min,
                    width_st_max,
                    length_st_mean,
                    length_st_median,
                    length_st_min,
                    length_st_max,
                    var_area_st_mean,
                    var_area_st_median,
                    var_area_st_min,
                    var_area_st_max,
                    var_width_st_mean,
                    var_width_st_median,
                    var_width_st_min,
                    var_width_st_max,
                    var_length_st_mean,
                    var_length_st_median,
                    var_length_st_min,
                    var_length_st_max,
                    guardCell_length_mean,
                    guardCell_length_median,
                    guardCell_length_min,
                    guardCell_length_max,
                    guardCell_width_mean,
                    guardCell_width_median,
                    guardCell_width_min,
                    guardCell_width_max,
                    guardCell_area_mean,
                    guardCell_area_median,
                    guardCell_area_min,
                    guardCell_area_max,
                    guardCell_angle_mean,
                    guardCell_angle_median,
                    guardCell_angle_min,
                    guardCell_angle_max,
                    var_width_guardCell_mean,
                    var_width_guardCell_median,
                    var_width_guardCell_min,
                    var_width_guardCell_max,
                    var_length_guardCell_mean,
                    var_length_guardCell_median,
                    var_length_guardCell_min,
                    var_length_guardCell_max,
                    var_area_guardCell_mean,
                    var_area_guardCell_median,
                    var_area_guardCell_min,
                    var_area_guardCell_max,
                    wst_density_mean,
                    wst_density_median,
                    wst_density_min,
                    wst_density_max,
                    ratio_area_st_gc_mean,
                    ratio_area_st_gc_median,
                    ratio_area_st_gc_min,
                    ratio_area_st_gc_max,
                    ratio_area_to_img_mean,
                    ratio_area_to_img_median,
                    ratio_area_to_img_min,
                    ratio_area_to_img_max,
                    var_angle_mean,
                    var_angle_median,
                    var_angle_min,
                    var_angle_max,
                    SEve_mean,
                    SEve_median,
                    SEve_min,
                    SEve_max,
                    SDiv_mean,
                    SDiv_median,
                    SDiv_min,
                    SDiv_max,
                    SAgg_mean,
                    SAgg_median,
                    SAgg_min,
                    SAgg_max], 
                    axis=1)
                
                Output_data = pd.DataFrame(data=Output_data)
                random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
                Output_data.to_excel(output_image_path_ + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
                self.show_info_messagebox_group_analysis()

        else:
            self.show_info_messagebox_group_analysis_2()
            pass          

    # #### Put a button to aqure the filedialog

    def loadInputFolder(self):
        """ """
        self.Inputfilefolder=str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Input Directory"))
        Input_path = self.lineEdit.setText(self.Inputfilefolder)

    def loadOutputFolder(self):
        """ """
        self.Outputfilefolder = str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Output Directory"))
        Output_path = self.lineEdit_2.setText(self.Outputfilefolder)

    def loadCsv_Box_model(self):
        """ """
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

    def loadCsv_Segment_model(self):
        """ """
        self.model = QtGui.QStandardItemModel(StoManager1)
        self.tableView.setModel(self.model)
        self.output_image_path_ = self.lineEdit_2.text()

        output_image_path_ = os.path.join(str(self.output_image_path_), "Predict_output/Output_csv")

        if bool(glob.glob(output_image_path_ + "/" + '*csv')) is True:

            for file_path in glob.glob(output_image_path_ + '/' + '*.csv'):
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
        """ """
        # clear the contents of the table view
        for i in range(self.model.rowCount()):
            for j in range(self.model.columnCount()):
                self.model.setData(self.model.index(i, j), "")
                
    def loadExl_Box_model(self):
        """ """
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

    def loadExl_Segment_model(self):
        """ """
        self.model = QtGui.QStandardItemModel(StoManager1)
        self.tableView.setModel(self.model)
        self.output_image_path_ = self.lineEdit_2.text()

        output_image_path_ = os.path.join(str(self.output_image_path_), "Predict_output/Output_csv")        
        
        if bool(glob.glob(output_image_path_ + "/" + '*csv')) is True:        
            if bool(glob.glob(output_image_path_ + '/' + '*.xlsx')):
                for file_path in glob.glob(output_image_path_ + '/' + '*.xlsx'):
                    newpath =  output_image_path_ + f'/New_folder' 
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
        """ """
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
        """ """
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
        """ """
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
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Opps... Now this function is only applicable for our Populus dataset.")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 Group analysis 📈")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        #msg.setInformativeText("I will add a customizable function for our users very soon 🐌.")
        
        # start the app
        x= msg.exec_() 

    def show_info_messagebox_group_analysis_2(self):
        """ """
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
        """ """
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
        """ """
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
        """ """
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
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Heads-up 🦉 You can scroll horizontally or vertically to view the whole image (●’◡’●).")
        # setting Message box window title
        msg.setWindowTitle("StoManager1 View images 👨‍🏫")
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        # msg.setInformativeText("I will add a customizable function for our users very soon 💕.")
        # start the app
        x= msg.exec_() 

    def messagebox_define_input_path(self):
        """ """
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
        """ """
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
        """ """
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


    def messagebox_if_input_path_contains_folder(self):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("The input path cannot contain subfolders")        
        # setting Message box window title
        msg.setWindowTitle("Check if input path contains subfolders 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please remove the subfolders, and try one more time. 🐻")       
        # start the app
        x = msg.exec_()      

    def statistical_options(self):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("The input path cannot contain subfolders")        
        # setting Message box window title
        msg.setWindowTitle("Check if input path contains subfolders 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        
        msg.setInformativeText("Please remove the subfolders, and try one more time. 🐻")       
        # start the app
        x = msg.exec_()         

    def closeExcel(self):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Cannot write the file, because the file is opening")        
        # setting Message box window title
        msg.setWindowTitle("Did you open excel or image file(s)? 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please close the opened Excel or Image files, and try one more time. 🐻")       
        # start the app
        x = msg.exec_()          

    def keyError(self):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("Some of your images have less than 4 observations to calculate")        
        # setting Message box window title
        msg.setWindowTitle("KeyError 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please go to output folder check the results. 🐻")       
        # start the app
        x = msg.exec_()    

    def runtimeWarning(self):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText("mean requires at least one data point")        
        # setting Message box window title
        msg.setWindowTitle("runtimeWarning 🐸")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please go to output folder check the results. 🐻")       
        # start the app
        x = msg.exec_()    

    def noStomata(self,e):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText(f''+ str(e) +" 🐸")        
        # setting Message box window title
        msg.setWindowTitle("No stomata detected")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("Please ensure that the input image has stomata. 🐻")       
        # start the app
        x = msg.exec_()    

    def outofIndex(self,e):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText(f''+ str(e) +" 🐸")        
        # setting Message box window title
        msg.setWindowTitle("Out of the range")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("This is the end of the range of the images. 🐻")       
        # start the app
        x = msg.exec_()

    def raiseError(self,e):
        """ """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setWindowIcon(QtGui.QIcon('StoManager.ico'))
        # setting message for Message Box
        msg.setText(f''+ str(e) +" 🐸")        
        # setting Message box window title
        msg.setWindowTitle("Error Raised")        
        # declaring buttons on Message Box
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.setInformativeText("There was something went wrong. None Type object error means no stoma detected in the given images. Division by zero means there is only one stoma detected. If it's GPU memory error, you can use the CPU version. 🐻")       
        # start the app
        x = msg.exec_()        

# webbrowser

    def web_link_github(self):
        """ """
        webbrowser.open('https://github.com/JiaxinWang123/StoManager.git')

    def web_link_google_scholar(self):
        """ """
        webbrowser.open("https://scholar.google.com/citations?user=7be6E64AAAAJ&hl=en")

    def web_link_arxiv(self):
        """ """
        webbrowser.open("https://arxiv.org/abs/2304.10450")   

    def web_link_Homepage(self):
        """ """
        webbrowser.open("https://www.jiaxinwang.us/")          

    def web_link_Early_versions(self):
        """ """
        webbrowser.open("https://zenodo.org/search?q=parent.id%3A7686022&f=allversions%3Atrue&l=list&p=1&s=10&sort=version")  
          

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    StoManager1 = QtWidgets.QMainWindow()
    ui = Ui_StoManager1()
    ui.setupUi(StoManager1)
    StoManager1.show()
    sys.exit(app.exec_())