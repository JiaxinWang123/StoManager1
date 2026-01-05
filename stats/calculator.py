"""Statistical calculations for stomata data."""

import pandas as pd
import numpy as np
import os
import glob
import string
import random
from typing import Callable, Optional, List

class StatisticsCalculator:
    """Calculates statistics from processed stomata data."""
    
    def __init__(self):
        pass
    
    def _remove_outliers(self, data, lower_percentile, upper_percentile):
        """Original outlier removal logic."""
        if len(data) == 0:
            return data
        Q1 = np.percentile(data, lower_percentile, method='midpoint')
        Q3 = np.percentile(data, upper_percentile, method='midpoint')
        IQR = Q3 - Q1
        upper_bound = Q3 + 1.5 * IQR
        lower_bound = Q1 - 1.5 * IQR
        return data[(data >= lower_bound) & (data <= upper_bound)]

    def _read_csv_robust(self, file_path):
        """Read CSV and handle missing headers by checking if first row is numeric."""
        try:
            df = pd.read_csv(file_path, low_memory=False)
            expected_cols = ['number_wst', 'area_wst', 'guardCell_area', 'Labels']
            has_expected = any(col in df.columns for col in expected_cols)
            
            if not has_expected:
                df_no_header = pd.read_csv(file_path, header=None, low_memory=False)
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
                if df_no_header.shape[1] == len(columns):
                    df_no_header.columns = columns
                    return df_no_header
            return df
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            return None

    def calculate_statistics_yolov8(self, output_path: str,
                                    progress_callback: Optional[Callable] = None):
        """Calculate statistics for YOLOv8 segmentation results."""
        csv_path = os.path.join(output_path, "Predict_output/Output_csv")
        csv_files = glob.glob(os.path.join(csv_path, "*.csv"))
        csv_files = [f for f in csv_files if "Statistics.csv" not in f]
        
        if not csv_files: return

        results = self._initialize_yolov8_results()
        for idx, file_path in enumerate(csv_files):
            df = self._read_csv_robust(file_path)
            if df is None or df.shape[0] < 4: continue
            
            file_name = os.path.basename(file_path)[:-4]
            results['Filename'].append(file_name)
            self._process_yolov8_metrics_logic(df, results)
            if progress_callback: progress_callback(int((idx + 1) / len(csv_files) * 100))

        if results['Filename']: self._export_results(results, csv_path)

    def calculate_group_statistics_yolov8(self, output_path: str,
                                         progress_callback: Optional[Callable] = None):
        """Calculate grouped statistics for YOLOv8."""
        csv_path = os.path.join(output_path, "Predict_output/Output_csv")
        csv_files = glob.glob(os.path.join(csv_path, "*.csv"))
        csv_files = [f for f in csv_files if "Statistics.csv" not in f]
        
        if not csv_files: return

        results = self._initialize_grouped_results()
        processed_count = 0
        for idx, file_path in enumerate(csv_files):
            df = self._read_csv_robust(file_path)
            if df is None or df.shape[0] < 4: continue
            
            file_name = os.path.basename(file_path)[:-4]
            split_name = file_name.split(",")
            if len(split_name) < 5: continue
            
            results['Site'].append(split_name[0])
            results['Block'].append(split_name[1])
            results['Clone'].append(split_name[2])
            results['Month'].append(split_name[3])
            results['Year'].append(split_name[4])
            results['Filename'].append(file_name)
            
            self._process_yolov8_metrics_logic(df, results)
            processed_count += 1
            if progress_callback: progress_callback(int((idx + 1) / len(csv_files) * 100))
        
        if processed_count > 0: self._export_results(results, csv_path)

    def calculate_statistics_yolov3(self, output_path: str,
                                    progress_callback: Optional[Callable] = None):
        """Calculate statistics for YOLOv3 box detection results."""
        csv_files = glob.glob(os.path.join(output_path, "*.csv"))
        csv_files = [f for f in csv_files if "Statistics.csv" not in f]
        
        if not csv_files: return

        results = {
            'Filename': [], 'WST_Number': [], 'WST_Area_Ratio': [], 'WST_Density': [],
            'WST_Area_median': [], 'WST_Area_max': [], 'WST_Area_min': [],
            'WST_Area_mean': [], 'WST_Area_var': [], 'WST_Area_std': [],
            'WST_Orientation': [], 'WST_Orientation_var': []
        }
        
        for idx, file_path in enumerate(csv_files):
            try:
                df = pd.read_csv(file_path, low_memory=False)
                if df.shape[0] < 4: continue
                
                results['Filename'].append(os.path.basename(file_path)[:-4])
                orientation = df["Orientation"]
                Q1 = np.percentile(orientation, 25, method='midpoint')
                Q3 = np.percentile(orientation, 75, method='midpoint')
                IQR = Q3 - Q1
                orientation_cleaned = orientation[(orientation >= (Q1 - 1.5*IQR)) & (orientation <= (Q3 + 1.5*IQR))]
                
                wst_areas = df[df['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
                
                results['WST_Number'].append(np.mean(df["Num_of_whole_stomata"]))
                results['WST_Area_Ratio'].append(np.mean(df['Whole_stomata_area_ratio']))
                results['WST_Density'].append(int(np.mean(df["Whole_stomata_density"])))
                results['WST_Area_median'].append(np.median(wst_areas))
                results['WST_Area_max'].append(max(wst_areas))
                results['WST_Area_min'].append(min(wst_areas))
                results['WST_Area_mean'].append(np.mean(wst_areas))
                results['WST_Area_var'].append(np.var(wst_areas))
                results['WST_Area_std'].append(np.std(wst_areas))
                results['WST_Orientation'].append(np.median(orientation_cleaned))
                results['WST_Orientation_var'].append(np.var(orientation_cleaned))
                
                if progress_callback: progress_callback(int((idx + 1) / len(csv_files) * 100))
            except Exception as e:
                print(f"Error in YOLOv3 stats: {e}")
                continue

        df_out = pd.DataFrame(results)
        random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
        df_out.to_excel(os.path.join(output_path, f"Stomata_output_{random_str}.xlsx"), index=False)

    def _initialize_yolov8_results(self) -> dict:
        metrics = ['No_wst', 'box_w_wst', 'box_h_wst', 'area_wst', 'width_wst', 'length_wst',
                  'var_area_wst', 'var_width_wst', 'var_length_wst',
                  'No_st', 'box_w_st', 'box_h_st', 'area_st', 'width_st', 'length_st',
                  'var_area_st', 'var_width_st', 'var_length_st',
                  'guardCell_length', 'guardCell_width', 'guardCell_area', 'guardCell_angle',
                  'var_width_guardCell', 'var_length_guardCell',
                  'wst_density', 'ratio_area_st_gc', 'ratio_area_to_img', 'var_angle',
                  'SEve', 'SDiv', 'SAgg']
        stats = ['mean', 'median', 'min', 'max']
        results = {'Filename': []}
        for metric in metrics:
            for stat in stats: results[f'{metric}_{stat}'] = []
        return results

    def _initialize_grouped_results(self) -> dict:
        results = self._initialize_yolov8_results()
        group_cols = ['Site', 'Block', 'Clone', 'Month', 'Year']
        new_results = {col: [] for col in group_cols}
        new_results.update(results)
        return new_results

    def _process_yolov8_metrics_logic(self, df: pd.DataFrame, results: dict):
        metric_configs = {
            'No_wst': ('number_wst', 5, 95), 'box_w_wst': ('box_w_wst', 5, 95),
            'box_h_wst': ('box_h_wst', 2.5, 97.5), 'area_wst': ('area_wst', 2.5, 97.5),
            'width_wst': ('width_wst', 2.5, 97.5), 'length_wst': ('length_wst', 2.5, 97.5),
            'No_st': ('number_st', 5, 95), 'box_w_st': ('box_w_st', 2.5, 97.5),
            'box_h_st': ('box_h_st', 2.5, 97.5), 'area_st': ('area_st', 2.5, 97.5),
            'width_st': ('width_st', 2.5, 97.5), 'length_st': ('length_st', 2.5, 97.5),
            'guardCell_length': ('guardCell_length', 2.5, 97.5), 'guardCell_width': ('guardCell_width', 2.5, 97.5),
            'guardCell_area': ('guardCell_area', 2.5, 97.5), 'guardCell_angle': ('guardCell_angle', 2.5, 97.5),
            'ratio_area_st_gc': ('ratio_area_st_to_gc', 2.5, 97.5),
        }

        for res_key, (csv_col, low, high) in metric_configs.items():
            if csv_col in df.columns:
                data = pd.to_numeric(df[csv_col], errors='coerce').dropna()
                cleaned = self._remove_outliers(data, low, high)
                if len(cleaned) > 0:
                    results[f'{res_key}_mean'].append(np.mean(cleaned))
                    results[f'{res_key}_median'].append(np.median(cleaned))
                    results[f'{res_key}_min'].append(np.min(cleaned))
                    results[f'{res_key}_max'].append(np.max(cleaned))
                else:
                    for stat in ['mean', 'median', 'min', 'max']: results[f'{res_key}_{stat}'].append(0)
            else:
                for stat in ['mean', 'median', 'min', 'max']: results[f'{res_key}_{stat}'].append(0)

        other_metrics = {
            'var_area_wst': 'var_area_wst', 'var_width_wst': 'var_width_wst', 'var_length_wst': 'var_length_wst',
            'var_area_st': 'var_area_st', 'var_width_st': 'var_width_st', 'var_length_st': 'var_length_st',
            'var_width_guardCell': 'var_width_guardCell', 'var_length_guardCell': 'var_length_guardCell',
            'wst_density': 'wst_density', 'ratio_area_to_img': 'ratio_area_to_img', 'var_angle': 'var_angle',
            'SEve': 'SEve', 'SDiv': 'SDiv', 'SAgg': 'SAgg'
        }

        for res_key, csv_col in other_metrics.items():
            if csv_col in df.columns:
                data = pd.to_numeric(df[csv_col], errors='coerce').dropna()
                if len(data) > 0:
                    results[f'{res_key}_mean'].append(np.mean(data))
                    results[f'{res_key}_median'].append(np.median(data))
                    results[f'{res_key}_min'].append(np.min(data))
                    results[f'{res_key}_max'].append(np.max(data))
                else:
                    for stat in ['mean', 'median', 'min', 'max']: results[f'{res_key}_{stat}'].append(0)
            else:
                for stat in ['mean', 'median', 'min', 'max']: results[f'{res_key}_{stat}'].append(0)

    def _export_results(self, results: dict, output_path: str):
        lengths = {k: len(v) for k, v in results.items()}
        max_len = max(lengths.values())
        for k, v in results.items():
            if len(v) < max_len: v.extend([0] * (max_len - len(v)))
        
        df = pd.DataFrame(results)
        for col in df.columns:
            if col in ['Filename', 'Site', 'Block', 'Clone', 'Month', 'Year']: continue
            try:
                if any(x in col for x in ['ratio_area_st_to_gc', 'ratio_area_to_img']):
                    df[col] = df[col].map('{:,.3f}'.format)
                elif any(x in col for x in ['SEve', 'SDiv', 'SAgg']):
                    df[col] = df[col].map('{:,.4f}'.format)
                else:
                    df[col] = df[col].map('{:,.0f}'.format)
            except: pass

        random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
        df.to_excel(os.path.join(output_path, f"Stomata_output_{random_str}.xlsx"), index=False)