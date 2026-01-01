from .common_imports import *

class TrainingWindow(QMainWindow):
    def __init__(self):
        super(TrainingWindow, self).__init__()
        self.setupUi(self)
        self.training_runner = None  # Store the training thread
        self.training_process = None  # Store the training subprocess
        self.training_running = False

    def setupUi(self, StoManager1):
        StoManager1.setObjectName("StoManager1")
        StoManager1.resize(1200, 800)
        StoManager1.setWindowIcon(QtGui.QIcon("assets/StoManager.ico"))
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
        # ... (rest of the setupUi code will be appended or I'll read more to ensure I have it all)

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
        StoManager1.setWindowTitle(_translate("TrainingWindow", "StoManager1_model_trainer"))
        self.lineEdit.setText(_translate("TrainingWindow", "Select training data input data.yaml"))
        self.pushButton_8.setText(_translate("TrainingWindow", "Input"))
        self.lineEdit_2.setText(_translate("TrainingWindow", "Select model_training_in_app.exe file"))
        self.pushButton_9.setText(_translate("TrainingWindow", "Trainer"))
        self.lineEdit_4.setText(_translate("TrainingWindow", "Number of training epochs, default is 1000"))
        self.lineEdit_5.setText(_translate("TrainingWindow", "Image size, default is 640"))
        self.lineEdit_6.setText(_translate("TrainingWindow", "Batch, 2-64 depends on your GPU or CPU, default is 2"))
        self.lineEdit_7.setText(_translate("TrainingWindow", "Fliplr, default is 0"))
        self.lineEdit_9.setText(_translate("TrainingWindow", "Workers, default is 0"))
        self.checkBox.setText(_translate("TrainingWindow", "AMP Check, uncheck it if you got nan for training metrics"))
        self.checkBox_2.setText(_translate("TrainingWindow", "Train on CPU, check it if you don't have a powerful GPU"))        
        self.pushButton_3.setText(_translate("TrainingWindow", "Start Training"))
        self.pushButton_6.setText(_translate("TrainingWindow", "Stop Training"))
        self.pushButton_7.setText(_translate("TrainingWindow", "Check training result"))
        self.label_3.setText(_translate("TrainingWindow", "© Jiaxin Wang.  For questions or requests 📧  coolwjx@foxmail.com; jw3994@msstate.edu"))
        self.label_9.setText(_translate("TrainingWindow", "💕   LU"))
        self.menuHelp.setTitle(_translate("TrainingWindow", "Help"))
        self.menuTraining.setTitle(_translate("TrainingWindow", "Training"))
        self.actionH.setText(_translate("TrainingWindow", "H"))
        self.action.setText(_translate("TrainingWindow", ")"))
        self.actionYOLOv8_seg_x.setText(_translate("TrainingWindow", "YOLOv8-seg-x"))
        self.actionGoogle_scholar.setText(_translate("TrainingWindow", "Google scholar"))
        self.actionarXiv.setText(_translate("TrainingWindow", "arXiv"))
        self.actionGitHub.setText(_translate("TrainingWindow", "GitHub"))
        self.actionMy_Homepage.setText(_translate("TrainingWindow", "Ultralytics"))
        self.actionEarly_versions.setText(_translate("TrainingWindow", "Early versions"))
        self.actionOnly_whole_stomata.setText(_translate("TrainingWindow", "Only whole_stomata"))
        self.actionOnly_stomata_aperture.setText(_translate("TrainingWindow", "Only stomata (aperture)"))
        self.actionOnly_guard_cell.setText(_translate("TrainingWindow", "Only guard cell"))
        self.actionAll_metrics.setText(_translate("TrainingWindow", "All metrics"))
        self.actionGroup_Analysis.setText(_translate("TrainingWindow", "Group Analysis"))
        self.actionOpen_training_window.setText(_translate("TrainingWindow", "Train YOLOv8-seg-x"))

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
            # Check if the script_path is an executable (.exe) or a python script
            if self.script_path.lower().endswith('.exe'):
                cmd = [self.script_path, self.data_input_path, self.model_output_path, str(self.epoch), str(self.imgsz),
                       str(self.batch), str(self.fliplr), str(self.workers), self.weights_path, self.amp, self.device]
            else:
                cmd = ['python', self.script_path, self.data_input_path, self.model_output_path, str(self.epoch), str(self.imgsz),
                       str(self.batch), str(self.fliplr), str(self.workers), self.weights_path, self.amp, self.device]

            process = subprocess.Popen(
                cmd,
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
