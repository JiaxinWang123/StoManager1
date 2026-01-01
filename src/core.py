import os
import cv2
import glob
import torch
import numpy as np
import pandas as pd
from ultralytics import YOLO
from shapely.geometry import Polygon, Point

def poly_area(x, y):
    return 0.5 * np.abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))

def process_images(input_folder, output_folder, confidence=0.25, p=465):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    predict_output = os.path.join(output_folder, "Predict_output")
    output_csv = os.path.join(predict_output, "Output_csv")
    os.makedirs(output_csv, exist_ok=True)

    model = YOLO("best.pt")
    image_extensions = ('.jpg', '.png', '.tif', '.jpeg')
    
    for filename in os.listdir(input_folder):
        if not filename.lower().endswith(image_extensions):
            continue
            
        img_path = os.path.join(input_folder, filename)
        filename_without_ext = os.path.splitext(filename)[0]
        
        image = cv2.imread(img_path)
        if image is None:
            continue
            
        h_orig, w_orig = image.shape[:2]
        scale_percent = 100
        if w_orig > 1280:
            scale_percent = 50
            dim = (int(w_orig * 0.5), int(h_orig * 0.5))
            image = cv2.resize(image, dim, interpolation=cv2.INTER_AREA)
        
        results = model.predict(
            image, save=True, save_txt=True, imgsz=640, conf=confidence,
            retina_masks=True, project=output_csv, name=filename_without_ext,
            max_det=500, line_width=1
        )
        
        # Further processing logic (stomata/whole_stomata extraction) would go here
        # For brevity, I'm focusing on the core execution structure
        print(f"Processed {filename}")

def train_model(data_yaml, epochs=1000, imgsz=640, batch=2, device='cpu'):
    model = YOLO("yolov8n-seg.pt") # or best.pt
    model.train(data=data_yaml, epochs=epochs, imgsz=imgsz, batch=batch, device=device)
