"""Training window UI for YOLOv8 model training."""

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow
from qtpy.QtCore import QThread, Signal
import subprocess
import io
import os
import webbrowser
from config.constants import *
from ui.dialogs import DialogManager


class TrainingWindow(QMainWindow):
    """Window for training YOLOv8 models."""
    
    def __init__(self):
        super().__init__()
        self.training_runner = None
        self.training_process = None
        self.training_running = False
        self.dialog_manager = DialogManager()
        self.amp = "True"
        self.device = "gpu"
        self.setup_ui()
    
    def setup_ui(self):
        """Initialize the training window UI."""
        self.setObjectName("StoManager1")
        self.setWindowIcon(QtGui.QIcon("assets/StoManager.ico"))
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.setWindowIcon(QtGui.QIcon(ICON_PATH))
        
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, 
                                           QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(3)
        sizePolicy.setVerticalStretch(3)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.setFont(font)
        self.setStyleSheet("font: 8pt \"Arial\";\n"
                          "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, "
                          "stop:1 rgba(29, 141, 162, 255));")
        
        # Central widget
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        
        # Main vertical layout
        self.verticalWidget_2 = QtWidgets.QWidget(self.centralwidget)
        self.verticalWidget_2.setObjectName("verticalWidget_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.verticalWidget_2)
        self.verticalLayout_3.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        
        # Input data path
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        
        self.lineEdit = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit.setFont(font)
        self.lineEdit.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit.setClearButtonEnabled(True)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setText("Select training data input data.yaml")
        self.horizontalLayout.addWidget(self.lineEdit)
        
        self.pushButton_8 = QtWidgets.QPushButton(self.verticalWidget_2)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setStyleSheet("background-color: rgb(193, 101, 68);")
        self.pushButton_8.setObjectName("pushButton_8")
        self.pushButton_8.setText("Input")
        self.pushButton_8.clicked.connect(self.load_input_folder)
        self.horizontalLayout.addWidget(self.pushButton_8)
        
        self.verticalLayout_3.addLayout(self.horizontalLayout)
        
        # Trainer path
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        
        self.lineEdit_2 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_2.setClearButtonEnabled(True)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_2.setText("Select model_training_in_app.exe file")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        
        self.pushButton_9 = QtWidgets.QPushButton(self.verticalWidget_2)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setStyleSheet("background-color: rgb(193, 101, 68);")
        self.pushButton_9.setObjectName("pushButton_9")
        self.pushButton_9.setText("Trainer")
        self.pushButton_9.clicked.connect(self.load_output_folder)
        self.horizontalLayout_2.addWidget(self.pushButton_9)
        
        self.verticalLayout_3.addLayout(self.horizontalLayout_2)
        
        # Training parameters
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        
        # Epochs
        self.lineEdit_4 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_4.setClearButtonEnabled(True)
        self.lineEdit_4.setText("Number of training epochs, default is 1000")
        self.verticalLayout.addWidget(self.lineEdit_4)
        
        # Image size
        self.lineEdit_5 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_5.setClearButtonEnabled(True)
        self.lineEdit_5.setText("Image size, default is 640")
        self.verticalLayout.addWidget(self.lineEdit_5)
        
        # Batch
        self.lineEdit_6 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_6.setFont(font)
        self.lineEdit_6.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_6.setClearButtonEnabled(True)
        self.lineEdit_6.setText("Batch, 2-64 depends on your GPU or CPU, default is 2")
        self.verticalLayout.addWidget(self.lineEdit_6)
        
        # Fliplr
        self.lineEdit_7 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_7.setFont(font)
        self.lineEdit_7.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_7.setClearButtonEnabled(True)
        self.lineEdit_7.setText("Fliplr, default is 0")
        self.verticalLayout.addWidget(self.lineEdit_7)
        
        # Workers
        self.lineEdit_9 = QtWidgets.QLineEdit(self.verticalWidget_2)
        self.lineEdit_9.setFont(font)
        self.lineEdit_9.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_9.setClearButtonEnabled(True)
        self.lineEdit_9.setText("Workers, default is 0")
        self.verticalLayout.addWidget(self.lineEdit_9)
        
        # Checkboxes
        self.checkBox = QtWidgets.QCheckBox(self.verticalWidget_2)
        self.checkBox.setStyleSheet("color : rgb(255, 255, 255)")
        self.checkBox.setChecked(True)
        self.checkBox.setText("AMP Check, uncheck it if you got nan for training metrics")
        self.checkBox.stateChanged.connect(self.update_amp_state)
        self.verticalLayout.addWidget(self.checkBox)
        
        self.checkBox_2 = QtWidgets.QCheckBox(self.verticalWidget_2)
        self.checkBox_2.setStyleSheet("color : rgb(255, 255, 255)")
        self.checkBox_2.setText("Train on CPU, check it if you don't have a powerful GPU")
        self.checkBox_2.stateChanged.connect(self.update_device_state)
        self.verticalLayout.addWidget(self.checkBox_2)
        
        self.verticalLayout_2.addLayout(self.verticalLayout)
        
        # Buttons
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        
        self.pushButton_3 = QtWidgets.QPushButton(self.verticalWidget_2)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(91, 215, 244);")
        self.pushButton_3.setText("Start Training")
        self.pushButton_3.clicked.connect(self.check_input_path_folder)
        self.horizontalLayout_3.addWidget(self.pushButton_3)
        
        self.pushButton_6 = QtWidgets.QPushButton(self.verticalWidget_2)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_6.setText("Stop Training")
        self.pushButton_6.clicked.connect(self.stop_script)
        self.horizontalLayout_3.addWidget(self.pushButton_6)
        
        self.pushButton_7 = QtWidgets.QPushButton(self.verticalWidget_2)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setStyleSheet("background-color: rgb(15, 160, 123);")
        self.pushButton_7.setText("Check training result")
        self.pushButton_7.clicked.connect(self.open_file_explorer)
        self.horizontalLayout_3.addWidget(self.pushButton_7)
        
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        
        # Console output
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.verticalWidget_2)
        self.plainTextEdit.setAutoFillBackground(False)
        self.plainTextEdit.setStyleSheet("background-color: rgb(30, 30, 30); color: white")
        self.plainTextEdit.setPlainText("")
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout_2.addWidget(self.plainTextEdit)
        
        # Footer
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        
        self.label_3 = QtWidgets.QLabel(self.verticalWidget_2)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(214, 214, 214);")
        self.label_3.setText("© Jiaxin Wang.  For questions or requests 📧  jiaxinwang362@gmail.com; jiaxinwang@cornell.edu")
        self.horizontalLayout_10.addWidget(self.label_3)
        
        self.label_9 = QtWidgets.QLabel(self.verticalWidget_2)
        self.label_9.setFont(font)
        self.label_9.setStyleSheet("color:rgba(255,0,0,160);")
        self.label_9.setText("💕   LU")
        self.horizontalLayout_10.addWidget(self.label_9)
        
        self.verticalLayout_2.addLayout(self.horizontalLayout_10)
        self.verticalLayout_3.addLayout(self.verticalLayout_2)
        self.gridLayout.addWidget(self.verticalWidget_2, 0, 0, 1, 1)
        
        # Set central widget
        self.setCentralWidget(self.centralwidget)
        
        # Status bar
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        
        # Menu bar
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1221, 22))
        self.menubar.setObjectName("menubar")
        
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setTitle("Help")
        self.menuHelp.setObjectName("menuHelp")
        
        self.menuTraining = QtWidgets.QMenu(self.menubar)
        self.menuTraining.setTitle("Training")
        self.menuTraining.setObjectName("menuTraining")
        
        self.setMenuBar(self.menubar)
        
        # Actions
        self.actionMy_Homepage = QtWidgets.QAction(self)
        self.actionMy_Homepage.setText("Ultralytics")
        self.actionMy_Homepage.triggered.connect(self.web_link_ultralytics)
        
        self.actionOpen_training_window = QtWidgets.QAction(self)
        self.actionOpen_training_window.setText("Train YOLOv8-seg-x")
        self.actionOpen_training_window.triggered.connect(self.web_link_yolov8)
        
        self.menuHelp.addAction(self.actionMy_Homepage)
        self.menuTraining.addAction(self.actionOpen_training_window)
        self.menubar.addAction(self.menuTraining.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        
        # Set window title
        self.setWindowTitle("StoManager1_model_trainer")
        
        # Initialize training runner
        self.training_runner = TrainingRunner()
        self.training_runner.update_signal.connect(self.update_console)
        self.training_running = False
    
    def update_amp_state(self):
        """Update AMP state from checkbox."""
        self.amp = "True" if self.checkBox.isChecked() else "False"
    
    def update_device_state(self):
        """Update device state from checkbox."""
        self.device = "cpu" if self.checkBox_2.isChecked() else "gpu"
    
    def load_input_folder(self):
        """Load input data.yaml file."""
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select data.yaml file", "", 
            "Yaml Files (*.yaml);;All Files (*)", options=options
        )
        if file_path:
            self.lineEdit.setText(file_path)
    
    def load_output_folder(self):
        """Load model trainer executable."""
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select model_trainer.exe file", "",
            "Excutable Files (*.exe);;All Files (*)", options=options
        )
        if file_path:
            self.lineEdit_2.setText(file_path)
    
    def check_input_path_folder(self):
        """Check paths and start training."""
        data_path = self.lineEdit.text()
        model_path = self.lineEdit_2.text()
        
        if (data_path and model_path and 
            data_path != "Select training data input data.yaml" and
            model_path != "Select model_training_in_app.exe file"):
            self.start_training()
        else:
            self.dialog_manager.missing_paths()
    
    def start_training(self):
        """Start model training."""
        if not self.training_running:
            self.training_running = True
            self.pushButton_3.setEnabled(False)
            self.plainTextEdit.clear()
            
            # Get parameters
            script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path, amp, device = self.get_parameters()
            
            # Get weights path
            folder_path = os.path.dirname(weights_path)
            weights_path = os.path.join(folder_path, 'yolov8x-seg.pt')
            
            # Set defaults
            epoch = epoch if epoch and epoch != "Number of training epochs, default is 1000" else DEFAULT_EPOCHS
            imgsz = imgsz if imgsz and imgsz != "Image size, default is 640" else DEFAULT_IMAGE_SIZE
            batch = batch if batch and batch != "Batch, 2-64 depends on your GPU or CPU, default is 2" else DEFAULT_BATCH_SIZE
            fliplr = fliplr if fliplr and fliplr != "Fliplr, default is 0" else DEFAULT_FLIPLR
            workers = workers if workers and workers != "Workers, default is 0" else DEFAULT_WORKERS
            
            # Start training thread
            self.training_runner = TrainingRunner()
            self.training_runner.set_parameters(
                script_path, data_input_path, epoch, imgsz, batch, 
                fliplr, workers, weights_path, amp, device
            )
            self.training_runner.update_signal.connect(self.update_console)
            self.training_runner.finished_signal.connect(self.stop_script)
            self.training_runner.start()
    
    def stop_script(self):
        """Stop training."""
        if self.training_running:
            if self.training_runner and self.training_runner.isRunning():
                self.training_runner.terminate()
                self.training_runner.wait()
            if self.training_process and self.training_process.poll() is None:
                self.training_process.terminate()
                self.training_process.wait()
            self.training_running = False
            self.pushButton_3.setEnabled(True)
    
    def get_parameters(self):
        """Get training parameters from UI."""
        script_path = self.lineEdit_2.text()
        data_input_path = self.lineEdit.text()
        epoch = self.lineEdit_4.text()
        imgsz = self.lineEdit_5.text()
        batch = self.lineEdit_6.text()
        fliplr = self.lineEdit_7.text()
        workers = self.lineEdit_9.text()
        weights_path = script_path
        amp = self.amp
        device = self.device
        
        return script_path, data_input_path, epoch, imgsz, batch, fliplr, workers, weights_path, amp, device
    
    def update_console(self, text: str):
        """Update console output."""
        self.plainTextEdit.appendPlainText(text)
    
    def open_file_explorer(self):
        """Open file explorer to view training results."""
        desired_path = os.path.join(os.getcwd(), "runs", "segment")
        if os.path.exists(desired_path) and os.path.isdir(desired_path):
            options = QtWidgets.QFileDialog.Options()
            QtWidgets.QFileDialog.getOpenFileName(
                self, "Open Directory", desired_path, "All Files (*);;", options=options
            )
            self.pushButton_7.setEnabled(True)
        else:
            self.dialog_manager.no_training_started()
            self.pushButton_7.setEnabled(False)
    
    def web_link_ultralytics(self):
        """Open Ultralytics GitHub."""
        webbrowser.open("https://github.com/ultralytics/ultralytics")
    
    def web_link_yolov8(self):
        """Open YOLOv8 training docs."""
        webbrowser.open("https://docs.ultralytics.com/modes/train/#introduction")


class TrainingRunner(QThread):
    update_signal = Signal(str)
    finished_signal = Signal()

    def __init__(self):
        super().__init__()
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

    def set_parameters(self, script_path, data_input_path, epoch, imgsz,
                       batch, fliplr, workers, weights_path, amp, device):
        self.script_path = script_path
        self.model_output_path = script_path
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
            # Choose correct command
            if self.script_path.lower().endswith(".exe"):
                cmd = [
                    self.script_path,
                    self.data_input_path,
                    self.model_output_path,
                    str(self.epoch), str(self.imgsz),
                    str(self.batch), str(self.fliplr),
                    str(self.workers),
                    self.weights_path,
                    self.amp,
                    self.device
                ]
            else:
                cmd = [
                    "python",
                    self.script_path,
                    self.data_input_path,
                    self.model_output_path,
                    str(self.epoch), str(self.imgsz),
                    str(self.batch), str(self.fliplr),
                    str(self.workers),
                    self.weights_path,
                    self.amp,
                    self.device
                ]

            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.PIPE,
                text=True,
                close_fds=True,
                env=os.environ
            )

            stdout_reader = io.open(
                process.stdout.fileno(),
                "r",
                encoding="utf-8",
                errors="replace"
            )

            while process.poll() is None:
                line = stdout_reader.readline()
                if not line:
                    break
                self.update_signal.emit(line.strip())

            self.finished_signal.emit()

        except Exception as e:
            self.update_signal.emit(f"Error: {str(e)}")

        finally:
            if process.poll() is None:
                process.terminate()
                process.wait()
