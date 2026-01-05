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
        """Process all images in folder with YOLOv3."""
        net = cv2.dnn.readNet(YOLOV3_WEIGHTS, YOLOV3_CONFIG)
        classes = STOMATA_CLASSES
        image_files = self.file_manager.get_image_files(input_path)
        
        if not image_files:
            raise ValueError("No image files found in input path")
        
        layer_names = net.getLayerNames()
        output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]
        colors = np.random.uniform(0, 100, size=(len(classes), 3))
        
        for img_num, img_path in enumerate(image_files, start=1):
            img = cv2.imread(img_path)
            if img is None: continue
            height, width, channels = img.shape
            
            blob = cv2.dnn.blobFromImage(img, 0.00392, (416, 416), (0, 0, 0), True, crop=False)
            net.setInput(blob)
            outs = net.forward(output_layers)
            
            class_ids, confidences, boxes = [], [], []
            number_of_whole_stomata, number_of_stomata = [], []
            list_of_width, list_of_height, list_of_x, list_of_y = [], [], [], []
            orientations, labels = [], []
            list_of_image_width, list_of_image_height = [], []
            list_of_all_stomata_areas, list_of_whole_stomatal_area = [], []
            class_ids_2 = []
            
            for out in outs:
                for detection in out:
                    scores = detection[5:]
                    class_id = np.argmax(scores)
                    conf = scores[class_id]
                    if conf > confidence:
                        center_x, center_y = detection[0] * width, detection[1] * height
                        w, h = detection[2] * width, detection[3] * height
                        boxes.append([center_x - w/2, center_y - h/2, w, h])
                        confidences.append(float(conf))
                        class_ids.append(class_id)
            
            indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.1, 0.4)
            for i in range(len(boxes)):
                if i in indexes:
                    x, y, w, h = boxes[i]
                    label = str(classes[class_ids[i]])
                    if label == "whole_stomata":
                        number_of_whole_stomata.append(class_ids[i])
                        area = ((w * h) * 0.6878 + 806) * (10000 / (pixel_size * pixel_size))
                        list_of_all_stomata_areas.append(area)
                        list_of_whole_stomatal_area.append(area)
                    elif label == "stomata":
                        number_of_stomata.append(class_ids[i])
                        area = ((w * h + 116.08) / 1.7684) * (10000 / (pixel_size * pixel_size))
                        list_of_all_stomata_areas.append(area)
                    
                    orientation = math.log(w / h) * (-92.2325) + 44.5222
                    orientations.append(orientation if orientation >= 0 else orientation + 180)
                    list_of_width.append(w); list_of_height.append(h)
                    list_of_x.append(x); list_of_y.append(y)
                    labels.append(label)
                    list_of_image_height.append(height); list_of_image_width.append(width)
                    class_ids_2.append(class_ids[i])
            
            try:
                if list_of_whole_stomatal_area:
                    ratio = sum(area / (10000 / (pixel_size**2)) for area in list_of_whole_stomatal_area) / (width * height)
                else: ratio = 0
                
                density = len(number_of_whole_stomata) / (height * width / (pixel_size * 10.0) ** 2) if height * width > 0 else 0
                
                image_data = {
                    "Labels": labels,
                    "Width_(pixels)": list_of_width,
                    "Height_(pixels)": list_of_height,
                    "Orientation": orientations,
                    "Num_of_Stomata": [len(number_of_stomata)] * len(labels),
                    "Num_of_whole_stomata": [len(number_of_whole_stomata)] * len(labels),
                    "Width_of_image_(pixels)": list_of_image_width,
                    "Height_of_image_(pixels)": list_of_image_height,
                    "All_sotmata_area_(mum2)": list_of_all_stomata_areas,
                    "Whole_stomata_area_ratio": [ratio] * len(labels),
                    "Whole_stomata_density": [density] * len(labels)
                }
                
                df = pd.DataFrame(image_data)
                csv_path = os.path.join(output_path, os.path.basename(img_path)[:-4] + ".csv")
                df.to_csv(csv_path, index=False) # Changed index=True to index=False
                
                # Save annotated image
                cv2.imwrite(os.path.join(output_path, os.path.basename(img_path)[:-4] + ".jpg"), img)
            except Exception as e:
                print(f"Error processing {img_path}: {e}")
            
            if progress_callback:
                progress_callback(int((img_num / len(image_files)) * 100))