"""Main application window - Part 1."""

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPixmap, QStandardItemModel
import webbrowser
import csv
import os
from config.constants import *
from ui.dialogs import DialogManager
from ui.training_window import TrainingWindow
from utils.file_utils import FileManager
from processors.yolov3_processor import YOLOv3Processor
from processors.yolov8_processor import YOLOv8Processor
from processors.file_normalizer import FileNormalizer
from stats.calculator import StatisticsCalculator


class MainWindow(QtWidgets.QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        # Initialize managers
        self.dialog_manager = DialogManager()
        self.file_manager = FileManager()
        self.yolov3_processor = YOLOv3Processor()
        self.yolov8_processor = YOLOv8Processor()
        self.file_normalizer = FileNormalizer()
        self.stats_calculator = StatisticsCalculator()
        
        # State
        self.selected_image_index = 0
        self.input_path = ""
        self.output_path = ""
        self.pixel_size = DEFAULT_PIXEL_SIZE
        self.confidence = DEFAULT_CONFIDENCE
        self.model = QStandardItemModel(self)
        
        self.setup_ui()
        self.connect_signals()
    
    def setup_ui(self):
        """Initialize the user interface."""
        self.setObjectName("StoManager1_v10_Seg-x_Hardwoods")
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
        self.gridLayout_3 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_3.setObjectName("gridLayout_3")
        
        # Main vertical layout
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setSpacing(7)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        
        # Main grid layout
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setContentsMargins(2, 2, 2, 2)
        self.gridLayout.setObjectName("gridLayout")
        
        # Create UI sections
        self._create_left_panel()
        self._create_center_panel()
        
        # Add to main layout
        self.verticalLayout_5.addLayout(self.gridLayout)
        
        # Footer
        self._create_footer()
        
        self.gridLayout_3.addLayout(self.verticalLayout_5, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        
        # Status bar
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        
        # Menu bar
        self._create_menu_bar()
        
        self.setWindowTitle(WINDOW_TITLE)
    
    def _create_left_panel(self):
        """Create left control panel."""
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        
        self.verticalLayout_8 = QtWidgets.QVBoxLayout()
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        
        # Input folder selection
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        
        font = QtGui.QFont("Arial", 8, QtGui.QFont.Normal)
        
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setFont(font)
        self.lineEdit.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit.setClearButtonEnabled(True)
        self.lineEdit.setText("Input path")
        self.horizontalLayout.addWidget(self.lineEdit)
        
        self.pushButton_8 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setStyleSheet("background-color: rgb(196, 124, 180);")
        self.pushButton_8.setText("Input")
        self.horizontalLayout.addWidget(self.pushButton_8)
        
        self.verticalLayout.addLayout(self.horizontalLayout)
        
        # Output folder selection
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_2.setClearButtonEnabled(True)
        self.lineEdit_2.setText("Output path")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setStyleSheet("background-color: rgb(196, 124, 180);")
        self.pushButton_9.setText("Output")
        self.horizontalLayout_2.addWidget(self.pushButton_9)
        
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        
        # Model selection
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        
        self.YOLOv8_seg_x = QtWidgets.QCheckBox(self.centralwidget)
        self.YOLOv8_seg_x.setFont(font)
        self.YOLOv8_seg_x.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.YOLOv8_seg_x.setChecked(True)
        self.YOLOv8_seg_x.setText("Segment model using trained YOLOv8-seg-x")
        self.horizontalLayout_5.addWidget(self.YOLOv8_seg_x)
        
        self.verticalLayout.addLayout(self.horizontalLayout_5)
        
        # Pixel size parameter
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_4.setClearButtonEnabled(True)
        self.lineEdit_4.setText("Input pixels in 0.1 mm, the default is: 465")
        self.verticalLayout.addWidget(self.lineEdit_4)
        
        # Confidence parameter
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.lineEdit_5.setClearButtonEnabled(True)
        self.lineEdit_5.setText("Confidence threshold for detection the default is 0.25")
        self.verticalLayout.addWidget(self.lineEdit_5)
        
        # Action buttons
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(79, 186, 112);")
        self.pushButton_3.setText("Start Process")
        self.horizontalLayout_3.addWidget(self.pushButton_3)
        
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_6.setText("Normalize File")
        self.horizontalLayout_3.addWidget(self.pushButton_6)
        
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setStyleSheet("background-color: rgb(189, 214, 67);")
        self.pushButton_7.setText("Statistical Analysis")
        self.horizontalLayout_3.addWidget(self.pushButton_7)
        
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.verticalLayout_8.addLayout(self.verticalLayout)
        
        # Progress bar
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setEnabled(True)
        self.progressBar.setStyleSheet("background-color: rgb(255, 255, 255);\n"
                                      "color: rgb(255, 255, 255);")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(True)
        self.verticalLayout_8.addWidget(self.progressBar)
        
        self.verticalLayout_4.addLayout(self.verticalLayout_8)
        
        # Data preview section
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        
        self.pushButton_15 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_15.setFont(font)
        self.pushButton_15.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_15.setText("Preview exported table for single image")
        self.horizontalLayout_9.addWidget(self.pushButton_15)
        
        self.pushButton_16 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_16.setFont(font)
        self.pushButton_16.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_16.setText("Preview exported statistics for all images")
        self.horizontalLayout_9.addWidget(self.pushButton_16)
        
        self.verticalLayout_2.addLayout(self.horizontalLayout_9)
        
        # Table view
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        self.tableView.setFont(font)
        self.tableView.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.tableView.setModel(self.model)
        self.verticalLayout_2.addWidget(self.tableView)
        
        self.verticalLayout_4.addLayout(self.verticalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout_4, 0, 0, 1, 1)
    
    def _create_center_panel(self):
        """Create center image viewing panel."""
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        
        font = QtGui.QFont("Arial", 8, QtGui.QFont.Normal)
        
        # Top image viewer (original images)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton.setText("<< View previous")
        self.horizontalLayout_6.addWidget(self.pushButton)
        
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_4.setText("View original images")
        self.horizontalLayout_6.addWidget(self.pushButton_4)
        
        self.pushButton_10 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_10.setFont(font)
        self.pushButton_10.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_10.setText("View next >>")
        self.horizontalLayout_6.addWidget(self.pushButton_10)
        
        self.verticalLayout_3.addLayout(self.horizontalLayout_6)
        
        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, 
                                          QtWidgets.QSizePolicy.Expanding)
        self.graphicsView.setSizePolicy(sizePolicy)
        self.graphicsView.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.verticalLayout_3.addWidget(self.graphicsView)
        
        self.gridLayout_2.addLayout(self.verticalLayout_3, 0, 0, 1, 1)
        
        # Bottom image viewer (processed images)
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_2.setText("<< View previous")
        self.horizontalLayout_7.addWidget(self.pushButton_2)
        
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_5.setText("View detected images")
        self.horizontalLayout_7.addWidget(self.pushButton_5)
        
        self.pushButton_11 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_11.setFont(font)
        self.pushButton_11.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_11.setText("View next >>")
        self.horizontalLayout_7.addWidget(self.pushButton_11)
        
        self.verticalLayout_7.addLayout(self.horizontalLayout_7)
        
        self.graphicsView_2 = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView_2.setEnabled(True)
        self.graphicsView_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.verticalLayout_7.addWidget(self.graphicsView_2)
        
        self.gridLayout_2.addLayout(self.verticalLayout_7, 1, 0, 1, 1)
        self.gridLayout.addLayout(self.gridLayout_2, 0, 2, 1, 1)
    
    def _create_footer(self):
        """Create footer labels."""
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        
        font = QtGui.QFont("Arial", 8, QtGui.QFont.Normal)
        
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(214, 214, 214);")
        self.label_3.setText("© Jiaxin Wang.  For questions or requests 📧  coolwjx@foxmail.com; jw3994@msstate.edu")
        self.horizontalLayout_10.addWidget(self.label_3)
        
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setFont(font)
        self.label_9.setStyleSheet("color:rgba(255,0,0,160);")
        self.label_9.setText("💕   LU")
        self.horizontalLayout_10.addWidget(self.label_9)
        
        self.verticalLayout_5.addLayout(self.horizontalLayout_10)

    def _create_menu_bar(self):
        """Create menu bar with actions."""
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1155, 22))
        
        # Analysis menu
        self.menuAnalysis = QtWidgets.QMenu(self.menubar)
        self.menuAnalysis.setTitle("Analysis")
        
        self.actionOnly_whole_stomata = QtWidgets.QAction(self)
        self.actionOnly_whole_stomata.setCheckable(True)
        self.actionOnly_whole_stomata.setText("Only whole_stomata")
        
        self.actionOnly_stomata_aperture = QtWidgets.QAction(self)
        self.actionOnly_stomata_aperture.setCheckable(True)
        self.actionOnly_stomata_aperture.setText("Only stomata (aperture)")
        
        self.actionOnly_guard_cell = QtWidgets.QAction(self)
        self.actionOnly_guard_cell.setCheckable(True)
        self.actionOnly_guard_cell.setText("Only guard cell")
        
        self.actionAll_metrics = QtWidgets.QAction(self)
        self.actionAll_metrics.setCheckable(True)
        self.actionAll_metrics.setChecked(True)
        self.actionAll_metrics.setText("All metrics")
        
        self.actionGroup_Analysis = QtWidgets.QAction(self)
        self.actionGroup_Analysis.setCheckable(True)
        self.actionGroup_Analysis.setText("Group Analysis")
        self.actionGroup_Analysis.triggered.connect(self.statistics_group)
        
        self.menuAnalysis.addAction(self.actionOnly_whole_stomata)
        self.menuAnalysis.addAction(self.actionOnly_stomata_aperture)
        self.menuAnalysis.addAction(self.actionOnly_guard_cell)
        self.menuAnalysis.addSeparator()
        self.menuAnalysis.addAction(self.actionAll_metrics)
        self.menuAnalysis.addSeparator()
        self.menuAnalysis.addAction(self.actionGroup_Analysis)
        
        # Training menu
        self.menuTraining = QtWidgets.QMenu(self.menubar)
        self.menuTraining.setTitle("Training")
        
        self.actionOpen_training_window = QtWidgets.QAction(self)
        self.actionOpen_training_window.setText("Train YOLOv8-seg-x")
        self.actionOpen_training_window.triggered.connect(self.open_training_window)
        
        self.menuTraining.addAction(self.actionOpen_training_window)
        
        # Help menu
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setTitle("Help")
        
        self.actionGoogle_scholar = QtWidgets.QAction(self)
        self.actionGoogle_scholar.setText("Google scholar")
        self.actionGoogle_scholar.triggered.connect(self.web_link_google_scholar)
        
        self.actionarXiv = QtWidgets.QAction(self)
        self.actionarXiv.setText("arXiv")
        self.actionarXiv.triggered.connect(self.web_link_arxiv)
        
        self.actionGitHub = QtWidgets.QAction(self)
        self.actionGitHub.setText("GitHub")
        self.actionGitHub.triggered.connect(self.web_link_github)
        
        self.actionMy_Homepage = QtWidgets.QAction(self)
        self.actionMy_Homepage.setText("My Homepage")
        self.actionMy_Homepage.triggered.connect(self.web_link_homepage)
        
        self.actionEarly_versions = QtWidgets.QAction(self)
        self.actionEarly_versions.setText("Early versions")
        self.actionEarly_versions.triggered.connect(self.web_link_early_versions)
        
        self.menuHelp.addAction(self.actionGoogle_scholar)
        self.menuHelp.addAction(self.actionarXiv)
        self.menuHelp.addAction(self.actionGitHub)
        self.menuHelp.addAction(self.actionMy_Homepage)
        self.menuHelp.addSeparator()
        self.menuHelp.addAction(self.actionEarly_versions)
        
        # Add menus to bar
        self.menubar.addAction(self.menuAnalysis.menuAction())
        self.menubar.addAction(self.menuTraining.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        
        self.setMenuBar(self.menubar)
    
    def connect_signals(self):
        """Connect all button signals to handlers."""
        # File selection
        self.pushButton_8.clicked.connect(self.load_input_folder)
        self.pushButton_9.clicked.connect(self.load_output_folder)
        
        # Image navigation - input
        self.pushButton.clicked.connect(self.input_previous_image_clicked)
        self.pushButton_4.clicked.connect(lambda: [self.check_input_path(), self.dialog_manager.view_image_info()])
        self.pushButton_10.clicked.connect(self.input_next_image_clicked)
        
        # Image navigation - output
        self.pushButton_2.clicked.connect(self.output_previous_img)
        self.pushButton_5.clicked.connect(lambda: [self.check_output_path(), self.dialog_manager.view_image_info()])
        self.pushButton_11.clicked.connect(self.output_next_img)
        
        # Processing
        self.pushButton_3.clicked.connect(self.check_input_path_run)
        self.pushButton_6.clicked.connect(lambda: [self.dialog_manager.normalize_info(), self.normalize_filenames()])
        self.pushButton_7.clicked.connect(self.statistics)
        
        # Data viewing
        self.pushButton_15.clicked.connect(lambda: [self.clear_model(), self.load_csv()])
        self.pushButton_16.clicked.connect(lambda: [self.clear_model(), self.load_excel()])
    
    # ==================== File Selection ====================
    
    def load_input_folder(self):
        """Load input folder."""
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Input Directory")
        if folder:
            self.input_path = folder
            self.lineEdit.setText(folder)
    
    def load_output_folder(self):
        """Load output folder."""
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if folder:
            self.output_path = folder
            self.lineEdit_2.setText(folder)
    
    # ==================== Image Navigation ====================
    
    def check_input_path(self):
        """Check and display first input image."""
        try:
            self.input_path = self.lineEdit.text()
            if not self.file_manager.has_image_files(self.input_path):
                self.dialog_manager.no_input_path()
                return
            
            image_files = self.file_manager.get_image_files(self.input_path)
            self.selected_image_index = 0
            self._display_image(image_files[0], self.graphicsView)
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def input_next_image_clicked(self):
        """Show next input image."""
        try:
            self.input_path = self.lineEdit.text()
            if not self.file_manager.has_image_files(self.input_path):
                self.dialog_manager.no_input_path()
                return
            
            image_files = self.file_manager.get_image_files(self.input_path)
            self.selected_image_index = (self.selected_image_index + 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView)
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def input_previous_image_clicked(self):
        """Show previous input image."""
        try:
            self.input_path = self.lineEdit.text()
            if not self.file_manager.has_image_files(self.input_path):
                self.dialog_manager.no_input_path()
                return
            
            image_files = self.file_manager.get_image_files(self.input_path)
            self.selected_image_index = (self.selected_image_index - 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView)
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def check_output_path(self):
        """Check and display first output image."""
        try:
            self.output_path = self.lineEdit_2.text()
            
            if self.YOLOv8_seg_x.isChecked():
                output_csv_path = self.file_manager.get_output_csv_path(self.output_path, "yolov8")
                folders = [f for f in os.listdir(output_csv_path) 
                          if os.path.isdir(os.path.join(output_csv_path, f))]
                
                if not folders:
                    self.dialog_manager.no_output_path()
                    return
                
                image_files = []
                for folder in folders:
                    folder_path = os.path.join(output_csv_path, folder)
                    image_files.extend(self.file_manager.get_image_files(folder_path))
            else:
                image_files = self.file_manager.get_image_files(self.output_path)
            
            if not image_files:
                self.dialog_manager.no_output_path()
                return
            
            self.selected_image_index = 0
            self._display_image(image_files[0], self.graphicsView_2)
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def output_next_img(self):
        """Show next output image."""
        try:
            if self.YOLOv8_seg_x.isChecked():
                self._output_next_yolov8()
            else:
                self._output_next_yolov3()
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def output_previous_img(self):
        """Show previous output image."""
        try:
            if self.YOLOv8_seg_x.isChecked():
                self._output_previous_yolov8()
            else:
                self._output_previous_yolov3()
        except Exception as e:
            self.dialog_manager.out_of_range(e)
    
    def _output_next_yolov8(self):
        """Navigate to next YOLOv8 output image."""
        output_csv_path = self.file_manager.get_output_csv_path(self.output_path, "yolov8")
        image_files = self._get_yolov8_images(output_csv_path)
        
        if image_files:
            self.selected_image_index = (self.selected_image_index + 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView_2)
    
    def _output_previous_yolov8(self):
        """Navigate to previous YOLOv8 output image."""
        output_csv_path = self.file_manager.get_output_csv_path(self.output_path, "yolov8")
        image_files = self._get_yolov8_images(output_csv_path)
        
        if image_files:
            self.selected_image_index = (self.selected_image_index - 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView_2)
    
    def _output_next_yolov3(self):
        """Navigate to next YOLOv3 output image."""
        image_files = self.file_manager.get_image_files(self.output_path)
        if image_files:
            self.selected_image_index = (self.selected_image_index + 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView_2)
    
    def _output_previous_yolov3(self):
        """Navigate to previous YOLOv3 output image."""
        image_files = self.file_manager.get_image_files(self.output_path)
        if image_files:
            self.selected_image_index = (self.selected_image_index - 1) % len(image_files)
            self._display_image(image_files[self.selected_image_index], self.graphicsView_2)
    
    def _get_yolov8_images(self, output_csv_path: str):
        """Get all YOLOv8 output images."""
        image_files = []
        folders = [f for f in os.listdir(output_csv_path) 
                  if os.path.isdir(os.path.join(output_csv_path, f))]
        
        for folder in folders:
            folder_path = os.path.join(output_csv_path, folder)
            image_files.extend(self.file_manager.get_image_files(folder_path))
        
        return image_files
    
    def _display_image(self, image_path: str, viewer: QtWidgets.QGraphicsView):
        """Display an image in a graphics view."""
        scene = QtWidgets.QGraphicsScene(self)
        pixmap = QPixmap(image_path)
        item = QtWidgets.QGraphicsPixmapItem(pixmap)
        scene.addItem(item)
        viewer.setScene(scene)
    
    # ==================== Processing ====================
    
    def check_input_path_run(self):
        """Check input path and start processing."""
        self.input_path = self.lineEdit.text()
        
        if not self.file_manager.has_image_files(self.input_path):
            self.dialog_manager.no_input_path_process()
            return
        
        if self.file_manager.has_subfolders(self.input_path):
            self.dialog_manager.has_subfolders()
            return
        
        self.analysis_s()
    
    def analysis_s(self):
        """Run stomatal detection and measurement."""
        if self.YOLOv8_seg_x.isChecked():
            self.guard_cell()
        else:
            self.box_model()
    
    def guard_cell(self):
        """Process with YOLOv8 segmentation model."""
        try:
            pixel_size = self._get_parameter_value(self.lineEdit_4, DEFAULT_PIXEL_SIZE)
            confidence = self._get_parameter_value(self.lineEdit_5, DEFAULT_CONFIDENCE)
            
            self.yolov8_processor.process_folder(
                self.input_path,
                self.output_path,
                pixel_size,
                confidence,
                self.update_progress
            )
            
            self.dialog_manager.process_complete()
        except Exception as e:
            self.dialog_manager.show_error("Processing Error", e)
    
    def box_model(self):
        """Process with YOLOv3 box detection model."""
        try:
            pixel_size = self._get_parameter_value(self.lineEdit_4, DEFAULT_PIXEL_SIZE)
            confidence = self._get_parameter_value(self.lineEdit_5, DEFAULT_CONFIDENCE)
            
            self.yolov3_processor.process_folder(
                self.input_path,
                self.output_path,
                pixel_size,
                confidence,
                self.update_progress
            )
            
            self.dialog_manager.process_complete()
        except Exception as e:
            self.dialog_manager.show_error("Processing Error", e)
    
    def _get_parameter_value(self, line_edit: QtWidgets.QLineEdit, default: float) -> float:
        """Get parameter value from line edit with default."""
        text = line_edit.text().strip()
        if not text or text.startswith("Input") or text.startswith("Confidence"):
            return default
        try:
            return float(text)
        except ValueError:
            return default
    
    def update_progress(self, value: int):
        """Update progress bar."""
        self.progressBar.setValue(value)
        QtWidgets.QApplication.processEvents()
    
    # ==================== Statistics ====================
    
    def statistics(self):
        """Calculate statistics."""
        if self.YOLOv8_seg_x.isChecked():
            self.stata()
        else:
            self.stomata_no_groups_analysis()
    
    def stata(self):
        """Calculate YOLOv8 statistics."""
        try:
            self.stats_calculator.calculate_statistics_yolov8(
                self.output_path,
                self.update_progress
            )
            self.dialog_manager.statistics_complete()
        except Exception as e:
            if "No CSV files" in str(e):
                self.dialog_manager.no_processing_done()
            else:
                self.dialog_manager.show_error("Statistics Error", e)
    
    def stomata_no_groups_analysis(self):
        """Calculate YOLOv3 statistics."""
        try:
            self.stats_calculator.calculate_statistics_yolov3(
                self.output_path,
                self.update_progress
            )
            self.dialog_manager.statistics_complete()
        except Exception as e:
            if "No CSV files" in str(e):
                self.dialog_manager.no_processing_done()
            else:
                self.dialog_manager.show_error("Statistics Error", e)
    
    def statistics_group(self):
        """Calculate grouped statistics."""
        if self.YOLOv8_seg_x.isChecked():
            self.stomata_group_analysis_seg()
        else:
            self.stomata_group_analysis()
    
    def stomata_group_analysis_seg(self):
        """Calculate YOLOv8 grouped statistics."""
        try:
            self.stats_calculator.calculate_group_statistics_yolov8(
                self.output_path,
                self.update_progress
            )
            self.dialog_manager.group_analysis_info()
        except Exception as e:
            self.dialog_manager.show_error("Group Analysis Error", e)
    
    def stomata_group_analysis(self):
        """Calculate YOLOv3 grouped statistics (placeholder)."""
        self.dialog_manager.group_analysis_info()
    
    # ==================== File Normalization ====================
    
    def normalize_filenames(self):
        """Normalize filenames in input folder."""
        if not self.file_manager.has_image_files(self.input_path):
            self.dialog_manager.normalize_no_path()
            return
        
        try:
            renamed, skipped = self.file_normalizer.normalize_folder(self.input_path)
            print(f"Renamed {renamed} files, skipped {skipped}")
        except Exception as e:
            self.dialog_manager.show_error("Normalization Error", e)
    
    # ==================== Data Preview ====================
    
    def load_csv(self):
        """Load CSV for preview."""
        if self.YOLOv8_seg_x.isChecked():
            self.load_csv_segment_model()
        else:
            self.load_csv_box_model()
    
    def load_excel(self):
        """Load Excel for preview."""
        if self.YOLOv8_seg_x.isChecked():
            self.load_excel_segment_model()
        else:
            self.load_excel_box_model()
    
    def load_csv_segment_model(self):
        """Load CSV files from YOLOv8 output."""
        output_csv_path = self.file_manager.get_output_csv_path(self.output_path, "yolov8")
        csv_files = self.file_manager.get_csv_files(output_csv_path)
        
        if not csv_files:
            self.dialog_manager.no_processing_done()
            return
        
        self.model.clear()
        for file_path in csv_files:
            with open(file_path, "r") as f:
                for row in csv.reader(f):
                    items = [QtGui.QStandardItem(field) for field in row]
                    self.model.appendRow(items)
    
    def load_csv_box_model(self):
        """Load CSV files from YOLOv3 output."""
        csv_files = self.file_manager.get_csv_files(self.output_path)
        
        if not csv_files:
            self.dialog_manager.no_processing_done()
            return
        
        self.model.clear()
        for file_path in csv_files:
            with open(file_path, "r") as f:
                for row in csv.reader(f):
                    items = [QtGui.QStandardItem(field) for field in row]
                    self.model.appendRow(items)
    
    def load_excel_segment_model(self):
        """Load Excel files from YOLOv8 output."""
        import pandas as pd
        
        output_csv_path = self.file_manager.get_output_csv_path(self.output_path, "yolov8")
        xlsx_files = self.file_manager.get_xlsx_files(output_csv_path)
        
        if not xlsx_files:
            self.dialog_manager.no_statistics_file()
            return
        
        self._load_excel_files(xlsx_files, output_csv_path)
    
    def load_excel_box_model(self):
        """Load Excel files from YOLOv3 output."""
        import pandas as pd
        
        xlsx_files = self.file_manager.get_xlsx_files(self.output_path)
        
        if not xlsx_files:
            self.dialog_manager.no_statistics_file()
            return
        
        self._load_excel_files(xlsx_files, self.output_path)
    
    def _load_excel_files(self, xlsx_files, base_path):
        """Helper to load Excel files into table view."""
        import pandas as pd
        import string
        import random
        
        temp_folder = os.path.join(base_path, 'New_folder')
        self.file_manager.create_directory(temp_folder)
        self.file_manager.clean_directory(temp_folder)
        
        self.model.clear()
        
        for file_path in xlsx_files:
            random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
            csv_path = os.path.join(temp_folder, f"{random_str}.csv")
            pd.read_excel(file_path).to_csv(csv_path, index=None, header=True)
            
            with open(csv_path, "r") as f:
                for row in csv.reader(f):
                    items = [QtGui.QStandardItem(field) for field in row]
                    self.model.appendRow(items)
    
    def clear_model(self):
        """Clear table view model."""
        self.model.clear()
    
    # ==================== Menu Actions ====================
    
    def open_training_window(self):
        """Open training window."""
        self.training_window = TrainingWindow()
        self.training_window.show()
    
    def web_link_github(self):
        """Open GitHub repository."""
        webbrowser.open('https://github.com/JiaxinWang123/StoManager.git')
    
    def web_link_google_scholar(self):
        """Open Google Scholar profile."""
        webbrowser.open("https://scholar.google.com/citations?user=7be6E64AAAAJ&hl=en")
    
    def web_link_arxiv(self):
        """Open arXiv paper."""
        webbrowser.open("https://arxiv.org/abs/2304.10450")
    
    def web_link_homepage(self):
        """Open homepage."""
        webbrowser.open("https://www.jiaxinwang.us/")
    
    def web_link_early_versions(self):
        """Open early versions."""
        webbrowser.open("https://zenodo.org/search?q=parent.id%3A7686022&f=allversions%3Atrue&l=list&p=1&s=10&sort=version")

