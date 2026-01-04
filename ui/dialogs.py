"""Dialog and message box utilities."""

from PyQt5 import QtWidgets, QtGui
from config.constants import ICON_PATH


class DialogManager:
    """Manages all dialog boxes and message displays."""
    
    @staticmethod
    def _create_message_box(icon_type, title: str, message: str, details: str = "") -> QtWidgets.QMessageBox:
        """Create a message box with common settings.
        
        Args:
            icon_type: QMessageBox icon type
            title: Window title
            message: Main message text
            details: Additional informative text
            
        Returns:
            Configured message box
        """
        msg = QtWidgets.QMessageBox()
        msg.setIcon(icon_type)
        msg.setWindowIcon(QtGui.QIcon(ICON_PATH))
        msg.setText(message)
        msg.setWindowTitle(title)
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        if details:
            msg.setInformativeText(details)
        return msg
    
    @classmethod
    def show_info(cls, title: str, message: str, details: str = "") -> None:
        """Show information message box."""
        msg = cls._create_message_box(QtWidgets.QMessageBox.Information, title, message, details)
        msg.exec_()
    
    @classmethod
    def show_warning(cls, title: str, message: str, details: str = "") -> None:
        """Show warning message box."""
        msg = cls._create_message_box(QtWidgets.QMessageBox.Warning, title, message, details)
        msg.exec_()
    
    # Processing messages
    @classmethod
    def process_complete(cls) -> None:
        cls.show_info(
            "StoManager1 Processor 🐸",
            "Yeah! Process finished!",
            "Now you can check your output or do statistical analysis. 🐻"
        )
    
    @classmethod
    def statistics_complete(cls) -> None:
        cls.show_info(
            "StoManager1 Statistical analysis 🐼",
            "Sweet 😘~  Now you can preview statistical results. 🐮",
            "💕"
        )
    
    # Path validation messages
    @classmethod
    def no_input_path(cls) -> None:
        cls.show_warning(
            "StoManager1 View Original images 🐸",
            "Oh... You may forget define input path....",
            "Please define input image path, and try one more time. 🐻"
        )
    
    @classmethod
    def no_output_path(cls) -> None:
        cls.show_warning(
            "StoManager1 View Detected images 🐸",
            "Oh... You may forget define output path....",
            "Please define output image path, and try one more time. 🐻"
        )
    
    @classmethod
    def no_input_path_process(cls) -> None:
        cls.show_warning(
            "StoManager1 Processor images 🐸",
            "Oh... You may forget define input path....",
            "Please define input image path, and try one more time. 🐻"
        )
    
    @classmethod
    def missing_paths(cls) -> None:
        cls.show_warning(
            "Define your data file and model trainer 🐸",
            "Oh... You may forget define path of training data and model trainer....",
            "Please define both file path, and try one more time. 🐻"
        )
    
    # Folder/file issues
    @classmethod
    def has_subfolders(cls) -> None:
        cls.show_warning(
            "Check if input path contains subfolders 🐸",
            "The input path cannot contain subfolders",
            "Please remove the subfolders, and try one more time. 🐻"
        )
    
    @classmethod
    def no_statistics_file(cls) -> None:
        cls.show_info(
            "StoManager1 Statistical results preview 📊",
            "Opps...  there is no Statistics file. 🦊",
            "Have you done Group or Non-group analysis? 👀"
        )
    
    @classmethod
    def no_processing_done(cls) -> None:
        cls.show_info(
            "StoManager1 Statistical analysis 📈",
            "Opps... It looks like you haven't done processing images.",
            "Please define file paths, start process, and give it one more try  🌏."
        )
    
    @classmethod
    def no_training_started(cls) -> None:
        cls.show_warning(
            "Define your data file and model trainer and start your training 🐸",
            "Oh... You haven't started training yet....",
            "Please define your data file and model trainer and start your training, and try one more time. 🐻"
        )
    
    # File operation errors
    @classmethod
    def file_open_error(cls) -> None:
        cls.show_warning(
            "Did you open excel or image file(s)? 🐸",
            "Cannot write the file, because the file is opening",
            "Please close the opened Excel or Image files, and try one more time. 🐻"
        )
    
    # Data issues
    @classmethod
    def key_error(cls) -> None:
        cls.show_warning(
            "KeyError 🐸",
            "Some of your images have less than 4 observations to calculate",
            "Please go to output folder check the results. 🐻"
        )
    
    @classmethod
    def runtime_warning(cls) -> None:
        cls.show_warning(
            "runtimeWarning 🐸",
            "mean requires at least one data point",
            "Please go to output folder check the results. 🐻"
        )
    
    @classmethod
    def out_of_range(cls, error: Exception) -> None:
        cls.show_warning(
            "Out of the range",
            f"{str(error)} 🐸",
            "This is the end of the range of the images. 🐻"
        )
    
    # Info messages
    @classmethod
    def view_image_info(cls) -> None:
        cls.show_info(
            "StoManager1 View images 👨‍🏫",
            "Heads-up 🦉 You can scroll horizontally or vertically to view the whole image (◕‿◕)."
        )
    
    @classmethod
    def normalize_info(cls) -> None:
        cls.show_info(
            "StoManager1 FileName Normalizer 📶",
            "Please note that now filename normalizer is only applicable for our Populus dataset. "
            "It will work if you are playing our Populus dataset.",
            "I will add a customizable function for our users very soon 💕."
        )
    
    @classmethod
    def normalize_no_path(cls) -> None:
        cls.show_info(
            "StoManager1 FileName Normalizer 📶",
            "Opps... It looks like you haven't define input file path.",
            "Please define and input path and try one more time 💕."
        )
    
    @classmethod
    def group_analysis_info(cls) -> None:
        cls.show_info(
            "StoManager1 Group analysis 📈",
            "Opps... Now this function is only applicable for our Populus dataset.",
            ""
        )
    
    # General error
    @classmethod
    def show_error(cls, title: str, error: Exception) -> None:
        cls.show_warning(
            title,
            f"{str(error)} 🐸",
            "There was something went wrong. None Type object error means no stoma detected. "
            "Division by zero means there is only one stoma detected. "
            "If it's GPU memory error, you can use the CPU version. 🐻"
        )
