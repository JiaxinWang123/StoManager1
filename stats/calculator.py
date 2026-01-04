"""Statistical calculations for stomata data."""

import pandas as pd
import numpy as np
import os
import glob
import string
import random
from typing import Callable, Optional, List
from utils.file_utils import FileManager
from processors.stomata_analyzer import StomataAnalyzer


class StatisticsCalculator:
    """Calculates statistics from processed stomata data."""
    
    def __init__(self):
        self.file_manager = FileManager()
        self.analyzer = StomataAnalyzer()
    
    def calculate_statistics_yolov8(self, output_path: str,
                                    progress_callback: Optional[Callable] = None):
        """Calculate statistics for YOLOv8 segmentation results.
        
        Args:
            output_path: Base output path
            progress_callback: Progress update callback
        """
        csv_path = self.file_manager.get_output_csv_path(output_path, "yolov8")
        csv_files = self.file_manager.get_csv_files(csv_path)
        
        if not csv_files:
            raise ValueError("No CSV files found for statistics calculation")
        
        # Remove existing statistics file
        stats_file = os.path.join(csv_path, 'Statistics.csv')
        self.file_manager.remove_file_if_exists(stats_file)
        
        # Initialize result lists
        results = self._initialize_yolov8_results()
        
        # Process each CSV file
        for idx, file_path in enumerate(csv_files):
            try:
                df = pd.read_csv(file_path, low_memory=False)
                
                # Skip if too few observations
                if df.shape[0] < 4:
                    continue
                
                # Extract filename
                file_name = os.path.basename(file_path)[:-4]
                results['Filename'].append(file_name)
                
                # Process each metric
                self._process_yolov8_metrics(df, results)
                
                if progress_callback:
                    progress_callback(int((idx + 1) / len(csv_files) * 100))
                    
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
        
        # Export results
        self._export_yolov8_results(results, csv_path)
    
    def calculate_statistics_yolov3(self, output_path: str,
                                    progress_callback: Optional[Callable] = None):
        """Calculate statistics for YOLOv3 box detection results."""
        csv_files = self.file_manager.get_csv_files(output_path)
        
        if not csv_files:
            raise ValueError("No CSV files found for statistics calculation")
        
        # Remove existing statistics file
        stats_file = os.path.join(output_path, 'Statistics.csv')
        self.file_manager.remove_file_if_exists(stats_file)
        
        # Initialize result lists
        results = {
            'Filename': [],
            'WST_Number': [],
            'WST_Area_Ratio': [],
            'WST_Density': [],
            'WST_Area_median': [],
            'WST_Area_max': [],
            'WST_Area_min': [],
            'WST_Area_mean': [],
            'WST_Area_var': [],
            'WST_Area_std': [],
            'WST_Orientation': [],
            'WST_Orientation_var': []
        }
        
        for idx, file_path in enumerate(csv_files):
            try:
                df = pd.read_csv(file_path, low_memory=False)
                
                if df.shape[0] < 4:
                    continue
                
                file_name = os.path.basename(file_path)[:-4]
                results['Filename'].append(file_name)
                
                # Extract data
                orientation = df["Orientation"]
                whole_stomata_areas = df[df['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
                
                # Remove outliers from orientation
                orientation = self.analyzer.remove_outliers(orientation, 25, 75)
                
                # Calculate statistics
                results['WST_Number'].append(np.mean(df["Num_of_whole_stomata"]))
                results['WST_Area_Ratio'].append(np.mean(df['Whole_stomata_area_ratio']))
                results['WST_Density'].append(int(np.mean(df["Whole_stomata_density"])))
                results['WST_Area_median'].append(np.median(whole_stomata_areas))
                results['WST_Area_max'].append(max(whole_stomata_areas))
                results['WST_Area_min'].append(min(whole_stomata_areas))
                results['WST_Area_mean'].append(np.mean(whole_stomata_areas))
                results['WST_Area_var'].append(np.var(whole_stomata_areas))
                results['WST_Area_std'].append(np.std(whole_stomata_areas))
                results['WST_Orientation'].append(np.median(orientation))
                results['WST_Orientation_var'].append(np.var(orientation))
                
                if progress_callback:
                    progress_callback(int((idx + 1) / len(csv_files) * 100))
                    
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
        
        # Export
        df = pd.DataFrame(results)
        random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
        output_file = os.path.join(output_path, f"Stomata_output_{random_str}.xlsx")
        df.to_excel(output_file)
    
    def calculate_group_statistics_yolov8(self, output_path: str,
                                         progress_callback: Optional[Callable] = None):
        """Calculate grouped statistics for YOLOv8 (Populus dataset specific)."""
        csv_path = self.file_manager.get_output_csv_path(output_path, "yolov8")
        csv_files = self.file_manager.get_csv_files(csv_path)
        
        if not csv_files:
            raise ValueError("No CSV files found")
        
        # Check if filenames have grouping structure
        sample_file = os.path.basename(csv_files[0])[:-4]
        split_name = sample_file.split(",")
        
        if len(split_name) < 5:
            raise ValueError("Filenames do not have group structure (need Site,Block,Clone,Month,Year)")
        
        # Initialize grouped results
        results = self._initialize_grouped_results()
        
        for idx, file_path in enumerate(csv_files):
            try:
                df = pd.read_csv(file_path, low_memory=False)
                
                if df.shape[0] < 4:
                    continue
                
                # Extract grouping variables from filename
                file_name = os.path.basename(file_path)[:-4]
                split_name = file_name.split(",")
                
                results['Site'].append(split_name[0])
                results['Block'].append(split_name[1])
                results['Clone'].append(split_name[2])
                results['Month'].append(split_name[3])
                results['Year'].append(split_name[4])
                results['Filename'].append(file_name)
                
                # Process metrics with outlier removal
                self._process_grouped_metrics(df, results)
                
                if progress_callback:
                    progress_callback(int((idx + 1) / len(csv_files) * 100))
                    
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
        
        # Export grouped results
        self._export_grouped_results(results, csv_path)
    
    def _initialize_yolov8_results(self) -> dict:
        """Initialize result dictionary for YOLOv8 statistics."""
        metrics = ['No_wst', 'box_w_wst', 'box_h_wst', 'area_wst', 'width_wst', 'length_wst',
                  'var_area_wst', 'var_width_wst', 'var_length_wst',
                  'No_st', 'box_w_st', 'box_h_st', 'area_st', 'width_st', 'length_st',
                  'var_area_st', 'var_width_st', 'var_length_st',
                  'guardCell_length', 'guardCell_width', 'guardCell_area', 'guardCell_angle',
                  'var_width_guardCell', 'var_length_guardCell', 'var_area_guardCell',
                  'wst_density', 'ratio_area_st_gc', 'ratio_area_to_img', 'var_angle',
                  'SEve', 'SDiv', 'SAgg']
        
        stats = ['mean', 'median', 'min', 'max']
        
        results = {'Filename': []}
        for metric in metrics:
            for stat in stats:
                results[f'{metric}_{stat}'] = []
        
        return results
    
    def _initialize_grouped_results(self) -> dict:
        """Initialize result dictionary for grouped statistics."""
        results = self._initialize_yolov8_results()
        results['Site'] = []
        results['Block'] = []
        results['Clone'] = []
        results['Month'] = []
        results['Year'] = []
        return results
    
    def _process_yolov8_metrics(self, df: pd.DataFrame, results: dict):
        """Process all YOLOv8 metrics from a dataframe."""
        # Define columns and their outlier removal settings
        metrics = {
            'number_wst': (5, 95),
            'box_w_wst': (5, 95),
            'box_h_wst': (2.5, 97.5),
            'area_wst': (2.5, 97.5),
            'width_wst': (2.5, 97.5),
            'length_wst': (2.5, 97.5),
            'number_st': (5, 95),
            'box_w_st': (2.5, 97.5),
            'box_h_st': (2.5, 97.5),
            'area_st': (2.5, 97.5),
            'width_st': (2.5, 97.5),
            'length_st': (2.5, 97.5),
            'guardCell_length': (2.5, 97.5),
            'guardCell_width': (2.5, 97.5),
            'guardCell_area': (2.5, 97.5),
            'guardCell_angle': (2.5, 97.5),
            'ratio_area_st_to_gc': (2.5, 97.5)
        }
        
        # Process metrics with outlier removal
        for metric, (lower, upper) in metrics.items():
            if metric in df.columns:
                data = df[metric]
                cleaned_data = self.analyzer.remove_outliers(data, lower, upper)
                stats = self.analyzer.calculate_statistics(cleaned_data)
                
                results[f'{metric}_mean'].append(stats['mean'])
                results[f'{metric}_median'].append(stats['median'])
                results[f'{metric}_min'].append(stats['min'])
                results[f'{metric}_max'].append(stats['max'])
        
        # Process variance metrics (take last value)
        var_metrics = ['var_area_wst', 'var_width_wst', 'var_length_wst',
                      'var_area_st', 'var_width_st', 'var_length_st',
                      'var_width_guardCell', 'var_length_guardCell', 'var_area_guardCell',
                      'var_angle']
        
        for metric in var_metrics:
            if metric in df.columns:
                value = df[metric].iloc[-1] if len(df[metric]) > 0 else 0
                results[f'{metric}_mean'].append(value)
                results[f'{metric}_median'].append(value)
                results[f'{metric}_min'].append(value)
                results[f'{metric}_max'].append(value)
        
        # Process spatial indices
        for metric in ['SEve', 'SDiv', 'SAgg']:
            if metric in df.columns:
                value = df[metric].iloc[-1] if len(df[metric]) > 0 else 0
                results[f'{metric}_mean'].append(value)
                results[f'{metric}_median'].append(value)
                results[f'{metric}_min'].append(value)
                results[f'{metric}_max'].append(value)
        
        # Process other metrics
        if 'wst_density' in df.columns:
            val = df['wst_density'].iloc[-1] if len(df['wst_density']) > 0 else 0
            results['wst_density_mean'].append(val)
            results['wst_density_median'].append(val)
            results['wst_density_min'].append(val)
            results['wst_density_max'].append(val)
        
        if 'ratio_area_to_img' in df.columns:
            val = df['ratio_area_to_img'].iloc[-1] if len(df['ratio_area_to_img']) > 0 else 0
            results['ratio_area_to_img_mean'].append(val)
            results['ratio_area_to_img_median'].append(val)
            results['ratio_area_to_img_min'].append(val)
            results['ratio_area_to_img_max'].append(val)
    
    def _process_grouped_metrics(self, df: pd.DataFrame, results: dict):
        """Process metrics for grouped analysis."""
        self._process_yolov8_metrics(df, results)
    
    def _export_yolov8_results(self, results: dict, output_path: str):
        """Export YOLOv8 statistics to Excel."""
        # Convert to DataFrame and format
        df = pd.DataFrame(results)
        
        # Format numeric columns
        for col in df.columns:
            if col != 'Filename':
                if 'ratio' in col or 'SEve' in col or 'SDiv' in col or 'SAgg' in col:
                    df[col] = pd.Series(df[col], dtype=pd.Float64Dtype()).map('{:,.4f}'.format)
                else:
                    df[col] = pd.Series(df[col], dtype=pd.Float64Dtype()).map('{:,.0f}'.format)
        
        # Save
        random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
        output_file = os.path.join(output_path, f"Stomata_output_{random_str}.xlsx")
        df.to_excel(output_file, index=False)
    
    def _export_grouped_results(self, results: dict, output_path: str):
        """Export grouped statistics to Excel."""
        self._export_yolov8_results(results, output_path)
