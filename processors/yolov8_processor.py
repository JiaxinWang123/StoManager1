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
        """Process all images with YOLOv8 segmentation.
        
        Args:
            input_path: Path to input images
            output_path: Path to save results
            pixel_size: Pixels per 0.1mm
            confidence: Detection confidence threshold
            progress_callback: Progress update callback
        """
        # Create output directory structure
        predict_output = os.path.join(output_path, "Predict_output")
        output_csv_path = os.path.join(predict_output, "Output_csv")
        self.file_manager.create_directory(output_csv_path)
        
        # Load model
        model = YOLO(YOLOV8_WEIGHTS)
        
        # Get image files
        image_files = [f for f in os.listdir(input_path) 
                      if f.lower().endswith(('.jpg', '.png', '.tif', '.jpeg'))]
        
        if not image_files:
            raise ValueError("No image files found in input path")
        
        for img_num, filename in enumerate(image_files, start=1):
            try:
                img_path = os.path.join(input_path, filename)
                filename_without_ext = os.path.splitext(filename)[0]
                
                # Process single image
                self._process_single_image(
                    img_path, filename_without_ext, output_csv_path,
                    model, pixel_size, confidence
                )
                
                # Clear GPU cache
                torch.cuda.empty_cache()
                
                # Update progress
                if progress_callback:
                    progress_callback(int((img_num / len(image_files)) * 100))
                    
            except Exception as e:
                print(f"Error processing {filename}: {e}")
    
    def _process_single_image(self, img_path: str, filename: str, 
                             output_path: str, model, pixel_size: float, 
                             confidence: float):
        """Process a single image with YOLOv8."""
        # Read and prepare image
        images = cv2.imread(img_path)
        ori_img_shape = (images.shape[1], images.shape[0])
        
        # Resize if too large
        images, scale_percent = self.image_utils.resize_if_large(images, max_width=1280)
        H, W, _ = images.shape
        
        # Run prediction
        results = model.predict(
            images, save=True, save_txt=True, imgsz=640, conf=confidence,
            save_crop=False, line_thickness=1, retina_masks=True,
            project=output_path, name=filename, max_det=500
        )
        
        # Initialize data structures
        stomata = {"class": [], "box_w": [], "box_h": [], "area": [], 
                   "centorid": [], "length": [], "width": [], "angle": [], "img_shape": []}
        whole_stomata = {"class": [], "box_w": [], "box_h": [], "area": [],
                        "centorid": [], "length": [], "width": [], "angle": [], "img_shape": []}
        
        # Process each detection
        if len(results[0].masks.xyn) > 0:
            for i in range(len(results[0].masks.xyn)):
                box = results[0].boxes[i]
                box_w = box.xywh.tolist()[0][2] * (100 / scale_percent)
                box_h = box.xywh.tolist()[0][3] * (100 / scale_percent)
                cls = str(int(box.cls.tolist()[0]))
                
                # Get mask coordinates
                x = (results[0].masks.xyn[i][:, 0] * W).astype("int")
                y = (results[0].masks.xyn[i][:, 1] * H).astype("int")
                
                # Calculate polygon properties
                poly = Polygon(zip(x, y))
                box_rect = poly.minimum_rotated_rectangle
                x1, y1 = box_rect.exterior.coords.xy
                
                # Get dimensions
                edge_length = (
                    Point(x1[0], y1[0]).distance(Point(x1[1], y1[1])) * (100 / scale_percent),
                    Point(x1[1], y1[1]).distance(Point(x1[2], y1[2])) * (100 / scale_percent)
                )
                length = int(max(edge_length))
                width = int(min(edge_length))
                
                seg_area = int(self.image_utils.calculate_polygon_area(x, y)) * \
                          ((100 / scale_percent) * (100 / scale_percent))
                centroid = (int(sum(x) / len(x)), int(sum(y) / len(y)))
                
                # Calculate angle
                angle = int(self._azimuth(box_rect))
                
                # Store data
                if cls == "1":  # whole_stomata
                    whole_stomata["class"].append(cls)
                    whole_stomata["box_w"].append(box_w)
                    whole_stomata["box_h"].append(box_h)
                    whole_stomata["area"].append(seg_area)
                    whole_stomata["centorid"].append(centroid)
                    whole_stomata["length"].append(length)
                    whole_stomata["width"].append(width)
                    whole_stomata["angle"].append(angle)
                    whole_stomata["img_shape"].append(ori_img_shape)
                else:  # stomata pore
                    stomata["class"].append(cls)
                    stomata["box_w"].append(box_w)
                    stomata["box_h"].append(box_h)
                    stomata["area"].append(seg_area)
                    stomata["centorid"].append(centroid)
                    stomata["length"].append(length)
                    stomata["width"].append(width)
                    stomata["angle"].append(angle)
                    stomata["img_shape"].append(ori_img_shape)
        
        # Match stomata with whole_stomata and calculate guard cell metrics
        results_data = self._match_and_calculate(
            whole_stomata, stomata, pixel_size, ori_img_shape, results
        )
        
        # Save to CSV
        if results_data:
            df = pd.DataFrame(results_data)
            csv_path = os.path.join(output_path, f"{filename}.csv")
            df.to_csv(csv_path, index=False)
    
    def _match_and_calculate(self, whole_stomata: dict, stomata: dict, 
                            pixel_size: float, ori_img_shape: tuple, results):
        """Match stomata pores with whole stomata and calculate guard cell metrics."""
        res_whole_stomata = list(map(list, whole_stomata.values()))
        res_stomata = list(map(list, stomata.values()))
        
        link_st_wst = []
        guard_cell_area = []
        guard_cell_length = []
        guard_cell_width = []
        guard_cell_angle = []
        area_wst = []
        area_st = []
        width_wst = []
        width_st = []
        length_wst = []
        length_st = []
        
        # Calculate spatial indices
        if len(res_whole_stomata[0]) > 1:
            whole_stomata_centroids = res_whole_stomata[4]
            img_area = ori_img_shape[0] * ori_img_shape[1]
            spatial_indices = self.analyzer.calculate_spatial_indices(
                whole_stomata_centroids, img_area
            )
        else:
            spatial_indices = {'SEve': 0.0, 'SDiv': 0.0, 'SAgg': 0.0}
        
        # Match pores with whole stomata
        for j in range(len(res_whole_stomata[0])):
            whole_stomata_centroid = res_whole_stomata[4][j]
            area_1 = int((res_whole_stomata[3][j]) / (100 * pixel_size * pixel_size) * 1000000)
            length_1 = res_whole_stomata[5][j]
            width_1 = res_whole_stomata[6][j]
            angle_1 = res_whole_stomata[7][j]
            
            for k in range(len(res_stomata[0])):
                stomata_centroid = res_stomata[4][k]
                area_0 = int((res_stomata[3][k]) / (100 * pixel_size * pixel_size) * 1000000)
                length_0 = res_stomata[5][k]
                width_0 = res_stomata[6][k]
                
                # Check if pore is inside whole stomata
                distance = math.hypot(
                    stomata_centroid[0] - whole_stomata_centroid[0],
                    stomata_centroid[1] - whole_stomata_centroid[1]
                )
                
                if (distance <= 30 and length_1 > length_0 and 
                    area_1 - area_0 > 0 and width_1 - width_0 > 0):
                    
                    # Calculate guard cell metrics
                    guard_cell_area_ = area_1 - area_0
                    guard_cell_length_ = length_1 / (pixel_size / 100)
                    guard_cell_width_ = 0.5 * ((width_1 - width_0) / (pixel_size / 100))
                    
                    guard_cell_width.append(guard_cell_width_)
                    guard_cell_length.append(guard_cell_length_)
                    guard_cell_area.append(guard_cell_area_)
                    guard_cell_angle.append(angle_1)
                    
                    area_wst.append(area_1)
                    area_st.append(area_0)
                    width_wst.append(width_1 / (pixel_size / 100))
                    width_st.append(width_0 / (pixel_size / 100))
                    length_wst.append(length_1 / (pixel_size / 100))
                    length_st.append(length_0 / (pixel_size / 100))
                    
                    # Calculate metrics
                    number_st = results[0].boxes.cls.tolist().count(0.0)
                    number_wst = results[0].boxes.cls.tolist().count(1.0)
                    img_area = ori_img_shape[0] * ori_img_shape[1]
                    
                    wst_density = int(number_wst * ((100 * pixel_size * pixel_size) / img_area))
                    ratio_area_st_gc = float(f"{area_0 / guard_cell_area_:.3f}") if guard_cell_area_ > 0 else 0
                    ratio_area_to_img = float(f"{sum(a / (10000 / (pixel_size * pixel_size)) for a in area_wst) / img_area:.3f}")
                    
                    # Append to results
                    link_st_wst.append([
                        ori_img_shape, "1", number_wst, j,
                        int(res_whole_stomata[1][j]), int(res_whole_stomata[2][j]),
                        area_1, int(width_1 / (pixel_size / 100)), int(length_1 / (pixel_size / 100)),
                        int(np.var(area_wst)) if area_wst else 0,
                        int(np.var(width_wst)) if width_wst else 0,
                        int(np.var(length_wst)) if length_wst else 0,
                        whole_stomata_centroid,
                        "0", number_st, k,
                        int(res_stomata[1][k]), int(res_stomata[2][k]),
                        area_0, int(width_0 / (pixel_size / 100)), int(length_0 / (pixel_size / 100)),
                        int(np.var(area_st)) if area_st else 0,
                        int(np.var(width_st)) if width_st else 0,
                        int(np.var(length_st)) if length_st else 0,
                        stomata_centroid,
                        guard_cell_length_, guard_cell_width_, guard_cell_area_,
                        angle_1,
                        int(np.var(guard_cell_angle)) if guard_cell_angle else 0,
                        int(np.var(guard_cell_width)) if guard_cell_width else 0,
                        int(np.var(guard_cell_length)) if guard_cell_length else 0,
                        int(np.var(guard_cell_area)) if guard_cell_area else 0,
                        wst_density, ratio_area_st_gc, ratio_area_to_img,
                        spatial_indices['SEve'], spatial_indices['SDiv'], spatial_indices['SAgg']
                    ])
        
        return link_st_wst
    
    def _azimuth(self, mrr):
        """Calculate azimuth angle of minimum rotated rectangle."""
        bbox = list(mrr.exterior.coords)
        axis1 = self._dist(bbox[0], bbox[3])
        axis2 = self._dist(bbox[0], bbox[1])
        
        if axis1 <= axis2:
            az = self._azimuth_between_points(bbox[0], bbox[1])
        else:
            az = self._azimuth_between_points(bbox[0], bbox[3])
        return az
    
    def _azimuth_between_points(self, point1, point2):
        """Calculate azimuth between two points (0-180 degrees)."""
        angle = np.arctan2(point2[0] - point1[0], point2[1] - point1[1])
        return np.degrees(angle) if angle > 0 else np.degrees(angle) + 180
    
    def _dist(self, a, b):
        """Calculate distance between two points."""
        return math.hypot(b[0] - a[0], b[1] - a[1])
