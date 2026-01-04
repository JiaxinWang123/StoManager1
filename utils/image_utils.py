"""Image processing utilities."""

import cv2
import numpy as np
from typing import Tuple


class ImageUtils:
    """Utility functions for image operations."""
    
    @staticmethod
    def resize_if_large(image: np.ndarray, max_width: int = 1280) -> Tuple[np.ndarray, int]:
        """Resize image if it's too large to prevent GPU memory issues.
        
        Args:
            image: Input image array
            max_width: Maximum width threshold
            
        Returns:
            Tuple of (resized_image, scale_percent)
        """
        if image.shape[1] > max_width:
            scale_percent = 50
            width = int(image.shape[1] * scale_percent / 100)
            height = int(image.shape[0] * scale_percent / 100)
            dim = (width, height)
            resized = cv2.resize(image, dim, interpolation=cv2.INTER_AREA)
            return resized, scale_percent
        else:
            return image, 100
    
    @staticmethod
    def calculate_polygon_area(x: np.ndarray, y: np.ndarray) -> float:
        """Calculate area of a polygon using the shoelace formula.
        
        Args:
            x: Array of x coordinates
            y: Array of y coordinates
            
        Returns:
            Area of the polygon
        """
        return 0.5 * np.abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))
    
    @staticmethod
    def draw_detection(image: np.ndarray, 
                      box: Tuple[int, int, int, int],
                      label: str,
                      confidence: float,
                      color: Tuple[int, int, int]) -> None:
        """Draw bounding box and label on image.
        
        Args:
            image: Image to draw on
            box: (x, y, w, h) bounding box
            label: Class label
            confidence: Detection confidence
            color: RGB color tuple
        """
        x, y, w, h = box
        font = cv2.FONT_HERSHEY_PLAIN
        
        cv2.rectangle(image, (int(x), int(y)), (int(x + w), int(y + h)), color, 1)
        cv2.putText(image, label, (int(x), int(y) + 42), font, 1, color, 1)
        cv2.putText(image, str(int(h)), (int(x), int(y) + 60), font, 1, color, 1)
        cv2.putText(image, str(int(w)), (int(x), int(y) + 28), font, 1, color, 1)
        cv2.putText(image, str(round(confidence, 2)), (int(x), int(y) + 12), font, 1, color, 1)
