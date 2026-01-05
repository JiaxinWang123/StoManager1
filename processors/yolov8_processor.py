"""YOLOv8 segmentation processor."""

import cv2
import numpy as np
import pandas as pd
import torch
import math
import os
from ultralytics import YOLO
from shapely.geometry import Polygon
from shapely import Point
from typing import Callable, Optional
from config.constants import YOLOV8_WEIGHTS
from utils.file_utils import FileManager
from utils.image_utils import ImageUtils
from processors.stomata_analyzer import StomataAnalyzer

class YOLOv8Processor:
    """Processes images using YOLOv8 segmentation model."""
    
    def __init__(self):
        self.file_manager = FileManager()
        self.image_utils = ImageUtils()
        self.analyzer = StomataAnalyzer()
    
    def process_folder(self, input_path: str, output_path: str,
                      pixel_size: float, confidence: float,
                      progress_callback: Optional[Callable] = None):
        """Process all images with YOLOv8 segmentation."""
        predict_output = os.path.join(output_path, "Predict_output")
        output_csv_path = os.path.join(predict_output, "Output_csv")
        self.file_manager.create_directory(output_csv_path)
        
        model = YOLO(YOLOV8_WEIGHTS)
        image_files = [f for f in os.listdir(input_path) 
                      if f.lower().endswith(('.jpg', '.png', '.tif', '.jpeg'))]
        
        if not image_files:
            raise ValueError("No image files found in input path")
        
        for img_num, filename in enumerate(image_files, start=1):
            try:
                img_path = os.path.join(input_path, filename)
                filename_without_ext = os.path.splitext(filename)[0]
                
                self._process_single_image(
                    img_path, filename_without_ext, output_csv_path,
                    model, pixel_size, confidence
                )
                
                torch.cuda.empty_cache()
                if progress_callback:
                    progress_callback(int((img_num / len(image_files)) * 100))
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    def _azimuth(self, mrr):
        """Calculate azimuth angle of minimum rotated rectangle."""
        bbox = list(mrr.exterior.coords)
        axis1 = math.hypot(bbox[0][0] - bbox[3][0], bbox[0][1] - bbox[3][1])
        axis2 = math.hypot(bbox[0][0] - bbox[1][0], bbox[0][1] - bbox[1][1])
        
        if axis1 <= axis2:
            point1, point2 = bbox[0], bbox[1]
        else:
            point1, point2 = bbox[0], bbox[3]
            
        angle = np.arctan2(point2[0] - point1[0], point2[1] - point1[1])
        az = np.degrees(angle) if angle > 0 else np.degrees(angle) + 180
        return az

    def _process_single_image(self, img_path: str, filename: str, 
                             output_path: str, model, pixel_size: float, 
                             confidence: float):
        """Process a single image with YOLOv8."""
        images = cv2.imread(img_path)
        if images is None: return
        
        ori_img_shape = (images.shape[1], images.shape[0])
        images_resized, scale_percent = self.image_utils.resize_if_large(images, max_width=1280)
        H, W, _ = images_resized.shape
        
        results = model.predict(
            images_resized, save=True, save_txt=True, imgsz=640, conf=confidence,
            save_crop=False, line_thickness=1, retina_masks=True,
            project=output_path, name=filename, max_det=500
        )
        
        stomata = {"class": [], "box_w": [], "box_h": [], "area": [], 
                   "centroid": [], "length": [], "width": [], "angle": []}
        whole_stomata = {"class": [], "box_w": [], "box_h": [], "area": [],
                        "centroid": [], "length": [], "width": [], "angle": []}
        
        if results[0].masks is not None and len(results[0].masks.xyn) > 0:
            for i in range(len(results[0].masks.xyn)):
                box = results[0].boxes[i]
                box_w = box.xywh.tolist()[0][2] * (100 / scale_percent)
                box_h = box.xywh.tolist()[0][3] * (100 / scale_percent)
                cls = str(int(box.cls.tolist()[0]))
                
                x = (results[0].masks.xyn[i][:, 0] * W).astype("int")
                y = (results[0].masks.xyn[i][:, 1] * H).astype("int")
                
                if len(x) < 3: continue
                
                poly = Polygon(zip(x, y))
                box_rect = poly.minimum_rotated_rectangle
                x_rect, y_rect = box_rect.exterior.coords.xy
                
                edge_length = (
                    math.hypot(x_rect[0]-x_rect[1], y_rect[0]-y_rect[1]) * (100 / scale_percent),
                    math.hypot(x_rect[1]-x_rect[2], y_rect[1]-y_rect[2]) * (100 / scale_percent)
                )
                length = max(edge_length)
                width = min(edge_length)
                seg_area = poly.area * ((100 / scale_percent) ** 2)
                centroid = (int(sum(x) / len(x)) * (100 / scale_percent), 
                           int(sum(y) / len(y)) * (100 / scale_percent))
                angle = self._azimuth(box_rect)
                
                target = whole_stomata if cls == "1" else stomata
                target["class"].append(cls)
                target["box_w"].append(box_w)
                target["box_h"].append(box_h)
                target["area"].append(seg_area)
                target["centroid"].append(centroid)
                target["length"].append(length)
                target["width"].append(width)
                target["angle"].append(angle)

        results_list = self._match_and_calculate(
            whole_stomata, stomata, pixel_size, ori_img_shape, results
        )
        
        if results_list:
            columns = [
                "img_shape", "labels_wst", "number_wst", "index_wst",
                "box_w_wst", "box_h_wst", "area_wst", "width_wst", "length_wst",
                "var_area_wst", "var_width_wst", "var_length_wst", "centroid_wst",
                "labels_st", "number_st", "index_st",
                "box_w_st", "box_h_st", "area_st", "width_st", "length_st",
                "var_area_st", "var_width_st", "var_length_st", "centroid_st",
                "guardCell_length", "guardCell_width", "guardCell_area", "guardCell_angle",
                "var_angle", "var_width_guardCell", "var_length_guardCell", "var_area_guardCell",
                "wst_density", "ratio_area_st_to_gc", "ratio_area_to_img",
                "SEve", "SDiv", "SAgg"
            ]
            df = pd.DataFrame(results_list, columns=columns)
            csv_path = os.path.join(output_path, f"{filename}.csv")
            df.to_csv(csv_path, index=False)

    def _match_and_calculate(self, whole_stomata: dict, stomata: dict, 
                            pixel_size: float, ori_img_shape: tuple, results):
        """Match stomata pores with whole stomata and calculate guard cell metrics."""
        link_st_wst = []
        num_wst = len(whole_stomata["class"])
        num_st = len(stomata["class"])
        
        if num_wst > 1:
            spatial_indices = self.analyzer.calculate_spatial_indices(
                whole_stomata["centroid"], ori_img_shape[0] * ori_img_shape[1]
            )
        else:
            spatial_indices = {'SEve': 0.0, 'SDiv': 0.0, 'SAgg': 0.0}
        
        area_wst_list, width_wst_list, length_wst_list = [], [], []
        area_st_list, width_st_list, length_st_list = [], [], []
        gc_area_list, gc_width_list, gc_length_list, gc_angle_list = [], [], [], []

        for j in range(num_wst):
            wst_centroid = whole_stomata["centroid"][j]
            # Convert area to mum2
            area_1 = int((whole_stomata["area"][j]) / (100 * pixel_size * pixel_size) * 1000000)
            length_1 = whole_stomata["length"][j]
            width_1 = whole_stomata["width"][j]
            angle_1 = whole_stomata["angle"][j]
            
            for k in range(num_st):
                st_centroid = stomata["centroid"][k]
                area_0 = int((stomata["area"][k]) / (100 * pixel_size * pixel_size) * 1000000)
                length_0 = stomata["length"][k]
                width_0 = stomata["width"][k]
                
                dist = math.hypot(st_centroid[0] - wst_centroid[0], st_centroid[1] - wst_centroid[1])
                
                if (dist <= 30 and length_1 > length_0 and area_1 > area_0 and width_1 > width_0):
                    gc_area = area_1 - area_0
                    gc_length = length_1 / (pixel_size / 100)
                    gc_width = 0.5 * ((width_1 - width_0) / (pixel_size / 100))
                    
                    area_wst_list.append(area_1)
                    width_wst_list.append(width_1 / (pixel_size / 100))
                    length_wst_list.append(length_1 / (pixel_size / 100))
                    area_st_list.append(area_0)
                    width_st_list.append(width_0 / (pixel_size / 100))
                    length_st_list.append(length_0 / (pixel_size / 100))
                    gc_area_list.append(gc_area)
                    gc_width_list.append(gc_width)
                    gc_length_list.append(gc_length)
                    gc_angle_list.append(angle_1)
                    
                    img_area = ori_img_shape[0] * ori_img_shape[1]
                    wst_density = int(num_wst * ((100 * pixel_size * pixel_size) / img_area))
                    ratio_st_gc = area_0 / gc_area if gc_area > 0 else 0
                    
                    # Corrected ratio_area_to_img calculation matching original StoManager1_v10_new.py
                    # Original: sum(area/(10000/(pixel*pixel))) / (width * height)
                    # Which is: sum(area * pixel * pixel / 10000) / (width * height)
                    ratio_wst_img = sum(a * (pixel_size**2) / 10000 for a in area_wst_list) / img_area
                    
                    link_st_wst.append([
                        ori_img_shape, "1", num_wst, j,
                        int(whole_stomata["box_w"][j]), int(whole_stomata["box_h"][j]),
                        area_1, int(width_1 / (pixel_size / 100)), int(length_1 / (pixel_size / 100)),
                        int(np.var(area_wst_list)), int(np.var(width_wst_list)), int(np.var(length_wst_list)),
                        wst_centroid,
                        "0", num_st, k,
                        int(stomata["box_w"][k]), int(stomata["box_h"][k]),
                        area_0, int(width_0 / (pixel_size / 100)), int(length_0 / (pixel_size / 100)),
                        int(np.var(area_st_list)), int(np.var(width_st_list)), int(np.var(length_st_list)),
                        st_centroid,
                        gc_length, gc_width, gc_area, angle_1,
                        int(np.var(gc_angle_list)), int(np.var(gc_width_list)), int(np.var(gc_length_list)), int(np.var(gc_area_list)),
                        wst_density, ratio_st_gc, ratio_wst_img,
                        spatial_indices['SEve'], spatial_indices['SDiv'], spatial_indices['SAgg']
                    ])
        return link_st_wst