"""YOLOv3 detection processor."""

import cv2
import numpy as np
import pandas as pd
import math
import os
from typing import Callable, Optional
from config.constants import YOLOV3_WEIGHTS, YOLOV3_CONFIG, STOMATA_CLASSES
from utils.file_utils import FileManager
from utils.image_utils import ImageUtils


class YOLOv3Processor:
    """Processes images using YOLOv3 model."""
    
    def __init__(self):
        self.file_manager = FileManager()
        self.image_utils = ImageUtils()
    
    def process_folder(self, input_path: str, output_path: str, 
                      pixel_size: float, confidence: float,
                      progress_callback: Optional[Callable] = None):
        """Process all images in folder with YOLOv3.
        
        Args:
            input_path: Path to input images
            output_path: Path to save results
            pixel_size: Pixels per 0.1mm
            confidence: Detection confidence threshold
            progress_callback: Progress update callback
        """
        # Load YOLO model
        net = cv2.dnn.readNet(YOLOV3_WEIGHTS, YOLOV3_CONFIG)
        classes = STOMATA_CLASSES
        
        # Get image files
        image_files = self.file_manager.get_image_files(input_path)
        
        if not image_files:
            raise ValueError("No image files found in input path")
        
        layer_names = net.getLayerNames()
        output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]
        colors = np.random.uniform(0, 100, size=(len(classes), 3))
        
        # Process each image
        for img_num, img_path in enumerate(image_files, start=1):
            img = cv2.imread(img_path)
            height, width, channels = img.shape
            
            # Detecting objects
            blob = cv2.dnn.blobFromImage(img, 0.00392, (416, 416), (0, 0, 0), True, crop=False)
            net.setInput(blob)
            outs = net.forward(output_layers)
            
            # Initialize lists
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
            
            # Process detections
            for out in outs:
                for detection in out:
                    scores = detection[5:]
                    class_id = np.argmax(scores)
                    conf = scores[class_id]
                    
                    if conf > confidence:
                        center_x = (detection[0] * width)
                        center_y = (detection[1] * height)
                        w = (detection[2] * width)
                        h = (detection[3] * height)
                        x = (center_x - w / 2)
                        y = (center_y - h / 2)
                        
                        boxes.append([x, y, w, h])
                        centers.append([center_x, center_y])
                        confidences.append(float(conf))
                        class_ids.append(class_id)
            
            indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.1, 0.4)
            
            font = cv2.FONT_HERSHEY_PLAIN
            for i in range(len(boxes)):
                if i in indexes:
                    x, y, w, h = boxes[i]
                    label = str(classes[class_ids[i]])
                    color = colors[class_ids[i]]
                    
                    cv2.rectangle(img, (int(x), int(y)), (int(x) + int(w), (int(y) + int(h)), color, 1))
                    cv2.putText(img, label, (int(x), int(y) + 42), font, 1, color, 1)
                    cv2.putText(img, str(int(h)), (int(x), int(y) + 60), font, 1, color, 1)
                    cv2.putText(img, str(int(w)), (int(x), int(y) + 28), font, 1, color, 1)
                    cv2.putText(img, str(round(confidences[i], 2)), (int(x), int(y) + 12), font, 1, color, 1)
                    
                    if label == "whole_stomata":
                        number_of_whole_stomata.append(class_ids[i])
                        list_of_all_stomata_areas.append(
                            ((w * h) * 0.6878 + 806) * (10000 / (pixel_size * pixel_size)))
                        list_of_whole_stomatal_area.append(
                            ((w * h) * 0.6878 + 806) * (10000 / (pixel_size * pixel_size)))
                    elif label == "stomata":
                        number_of_stomata.append(class_ids[i])
                        list_of_all_stomata_areas.append(
                            ((w * h + 116.08) / 1.7684) * (10000 / (pixel_size * pixel_size)))
                    
                    orientation = math.log(w / h) * (-92.2325) + 44.5222
                    if orientation >= 0:
                        orientations.append(orientation)
                    else:
                        orientations.append(orientation + 180)
                    
                    list_of_width.append(w)
                    list_of_height.append(h)
                    list_of_x.append(x)
                    list_of_y.append(y)
                    image_paths.append(img_path)
                    labels.append(label)
                    list_of_image_height.append(height)
                    list_of_image_width.append(width)
                    class_ids_2.append(class_ids[i])
            
            # Save detection results to txt
            txt_path = os.path.join(output_path, os.path.basename(img_path)[:-4] + ".txt")
            with open(txt_path, 'w') as f:
                for m in range(len(list_of_width)):
                    w_2 = list_of_width[m] / width
                    h_2 = list_of_height[m] / height
                    x_center = (list_of_x[m] + list_of_width[m] / 2) / width
                    y_center = (list_of_y[m] + list_of_height[m] / 2) / height
                    f.write(str(class_ids_2[m]) + " " + str(confidences[m]) + " " + 
                           str(x_center) + " " + str(y_center) + " " + 
                           str(w_2) + " " + str(h_2) + "\n")
            
            # Calculate statistics
            try:
                if list_of_whole_stomatal_area:
                    list_of_whole_stomatal_area_ratio = sum(
                        (area / (10000 / (pixel_size * pixel_size))) 
                        for area in list_of_whole_stomatal_area
                    ) / (np.mean(list_of_image_width) * np.mean(list_of_image_height))
                else:
                    list_of_whole_stomatal_area_ratio = 0
                
                heights = list_of_height
                widths = list_of_width
                image_width = list_of_image_width
                image_height = list_of_image_height
                all_stomata_areas = [float(k) for k in list_of_all_stomata_areas]
                
                Whole_stomata_density = len(number_of_whole_stomata) / (
                    height * width / (pixel_size * 10.0) ** 2) if height * width > 0 else 0
                
                image_data = {
                    "Labels": labels,
                    "Width_(pixels)": widths,
                    "Height_(pixels)": heights,
                    "Orientation": orientations,
                    "Num_of_Stomata": len(number_of_stomata),
                    "Num_of_whole_stomata": len(number_of_whole_stomata),
                    "Width_of_image_(pixels)": image_width,
                    "Height_of_image_(pixels)": image_height,
                    "All_sotmata_area_(mum2)": all_stomata_areas,
                    "Whole_stomata_area_ratio": list_of_whole_stomatal_area_ratio,
                    "Whole_stomata_density": Whole_stomata_density
                }
                
                df = pd.DataFrame(image_data)
                csv_path = os.path.join(output_path, os.path.basename(img_path)[:-4] + ".csv")
                df.to_csv(csv_path, sep=',', index=True)
                
                # Save annotated image
                img_save_path = os.path.join(output_path, os.path.basename(img_path)[:-4] + ".jpg")
                cv2.imwrite(img_save_path, img)
                
            except Exception as e:
                print(f"Error processing {img_path}: {e}")
            
            # Update progress
            if progress_callback:
                progress_callback(int((img_num / len(image_files)) * 100))
