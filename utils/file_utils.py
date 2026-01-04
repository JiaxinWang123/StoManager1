"""File utility functions."""

import os
import glob
import shutil
from typing import List, Optional
from config.constants import ALLOWED_IMAGE_EXTENSIONS


class FileManager:
    """Handles file and directory operations."""
    
    @staticmethod
    def get_image_files(folder_path: str, extensions: tuple = ALLOWED_IMAGE_EXTENSIONS) -> List[str]:
        """Get all image files from a folder.
        
        Args:
            folder_path: Path to folder containing images
            extensions: Tuple of allowed file extensions
            
        Returns:
            List of image file paths
        """
        image_files = []
        for ext in extensions:
            pattern = os.path.join(folder_path, f'*.{ext}')
            image_files.extend(glob.glob(pattern))
        return image_files
    
    @staticmethod
    def has_image_files(folder_path: str) -> bool:
        """Check if folder contains image files.
        
        Args:
            folder_path: Path to check
            
        Returns:
            True if folder contains images
        """
        if not folder_path or not os.path.exists(folder_path):
            return False
        
        for ext in ALLOWED_IMAGE_EXTENSIONS:
            if glob.glob(os.path.join(folder_path, f"*{ext}")):
                return True
        return False
    
    @staticmethod
    def create_directory(path: str) -> str:
        """Create directory if it doesn't exist.
        
        Args:
            path: Directory path to create
            
        Returns:
            Created directory path
        """
        os.makedirs(path, exist_ok=True)
        return path
    
    @staticmethod
    def has_subfolders(folder_path: str) -> bool:
        """Check if a folder contains any subfolders.
        
        Args:
            folder_path: Path to check
            
        Returns:
            True if subfolders exist
        """
        if not os.path.exists(folder_path):
            return False
            
        for item in os.listdir(folder_path):
            if os.path.isdir(os.path.join(folder_path, item)):
                return True
        return False
    
    @staticmethod
    def get_csv_files(folder_path: str) -> List[str]:
        """Get all CSV files from a folder.
        
        Args:
            folder_path: Path to folder
            
        Returns:
            List of CSV file paths
        """
        return glob.glob(os.path.join(folder_path, '*.csv'))
    
    @staticmethod
    def get_xlsx_files(folder_path: str) -> List[str]:
        """Get all Excel files from a folder.
        
        Args:
            folder_path: Path to folder
            
        Returns:
            List of Excel file paths
        """
        return glob.glob(os.path.join(folder_path, '*.xlsx'))
    
    @staticmethod
    def remove_file_if_exists(file_path: str) -> None:
        """Remove a file if it exists.
        
        Args:
            file_path: Path to file
        """
        if os.path.exists(file_path):
            os.remove(file_path)
    
    @staticmethod
    def clean_directory(directory: str) -> None:
        """Remove all files in a directory.
        
        Args:
            directory: Directory to clean
        """
        if not os.path.exists(directory):
            return
            
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')
    
    @staticmethod
    def get_output_csv_path(output_folder: str, model_type: str = "yolov8") -> str:
        """Get the output CSV path based on model type.
        
        Args:
            output_folder: Base output folder
            model_type: Type of model ('yolov8' or 'yolov3')
            
        Returns:
            Path to CSV output directory
        """
        if model_type == "yolov8":
            return os.path.join(output_folder, "Predict_output", "Output_csv")
        else:
            return output_folder
