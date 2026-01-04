"""Stomata analysis and measurement processing."""

import numpy as np
import pandas as pd
from typing import Dict, List, Tuple, Optional
from scipy.spatial import distance, distance_matrix
from scipy.sparse.csgraph import minimum_spanning_tree
from config.constants import OUTLIER_LOWER_PERCENTILE, OUTLIER_UPPER_PERCENTILE, OUTLIER_IQR_MULTIPLIER


class StomataAnalyzer:
    """Analyzes stomata measurements and calculates statistics."""
    
    @staticmethod
    def remove_outliers(data: pd.Series, 
                       lower_percentile: float = OUTLIER_LOWER_PERCENTILE, 
                       upper_percentile: float = OUTLIER_UPPER_PERCENTILE,
                       iqr_multiplier: float = OUTLIER_IQR_MULTIPLIER) -> pd.Series:
        """Remove outliers from data using IQR method.
        
        Args:
            data: Pandas Series of data
            lower_percentile: Lower percentile threshold
            upper_percentile: Upper percentile threshold
            iqr_multiplier: IQR multiplier for bounds
            
        Returns:
            Data with outliers removed
        """
        Q1 = np.percentile(data, lower_percentile, method='midpoint')
        Q3 = np.percentile(data, upper_percentile, method='midpoint')
        IQR = Q3 - Q1
        
        upper_bound = Q3 + iqr_multiplier * IQR
        lower_bound = Q1 - iqr_multiplier * IQR
        
        upper_outliers = np.where(data >= upper_bound)
        lower_outliers = np.where(data <= lower_bound)
        
        data = data.copy()
        data.drop(upper_outliers[0], inplace=True, errors='ignore')
        data.drop(lower_outliers[0], inplace=True, errors='ignore')
        
        return data
    
    @staticmethod
    def calculate_spatial_indices(centroids: List[Tuple[float, float]], 
                                  img_area: float) -> Dict[str, float]:
        """Calculate stomata spatial distribution indices.
        
        Args:
            centroids: List of (x, y) centroid coordinates
            img_area: Total image area in pixels
            
        Returns:
            Dictionary with SEve, SDiv, SAgg values
        """
        centroids_array = np.array(centroids)
        n_stomata = len(centroids)
        
        if n_stomata < 2:
            return {'SEve': 0.0, 'SDiv': 0.0, 'SAgg': 0.0}
        
        # Calculate evenness (SEve)
        dist_matrix_vals = distance_matrix(centroids_array, centroids_array)
        mst = minimum_spanning_tree(dist_matrix_vals).toarray()
        PD = np.sum(mst, axis=1) / np.sum(mst)
        constant = 1 / (n_stomata - 1)
        evenness = (np.sum(PD[PD < constant]) + 
                   (n_stomata - 1 - len(PD[PD < constant])) * constant - constant) / (1 - constant)
        
        # Calculate divergence (SDiv)
        gravity_center = np.mean(centroids_array, axis=0)
        distances_to_gravity = np.array([
            distance.euclidean(centroids_array[i], gravity_center) 
            for i in range(n_stomata)
        ])
        mean_distance = np.mean(distances_to_gravity)
        sum_deviance = np.sum(distances_to_gravity - mean_distance)
        sum_abs_deviance = np.sum(np.abs(distances_to_gravity - mean_distance))
        divergence = (sum_deviance + mean_distance) / (sum_abs_deviance + mean_distance)
        
        # Calculate aggregation (SAgg)
        stomatal_density = n_stomata / img_area
        theoretical_distance = 1 / (2 * (stomatal_density ** 0.5)) if stomatal_density > 0 else 1
        nearest_neighbor_distances = np.array([
            np.sort(dist_matrix_vals[i])[1] for i in range(n_stomata)
        ])
        observed_distance = np.sum(nearest_neighbor_distances) / n_stomata
        aggregation = observed_distance / theoretical_distance if theoretical_distance > 0 else 0
        
        return {
            'SEve': float(f"{evenness:.4f}"),
            'SDiv': float(f"{divergence:.4f}"),
            'SAgg': float(f"{aggregation:.4f}")
        }
    
    @staticmethod
    def calculate_statistics(data: pd.Series) -> Dict[str, float]:
        """Calculate comprehensive statistics for a data series.
        
        Args:
            data: Pandas Series of measurements
            
        Returns:
            Dictionary with mean, median, min, max
        """
        if len(data) == 0:
            return {'mean': 0, 'median': 0, 'min': 0, 'max': 0}
            
        return {
            'mean': float(np.mean(data)),
            'median': float(np.median(data)),
            'min': float(min(data)),
            'max': float(max(data))
        }
    
    @staticmethod
    def calculate_guard_cell_metrics(area_wst: float, area_st: float,
                                    length_wst: float, width_wst: float,
                                    width_st: float, pixel_size: float) -> Dict[str, float]:
        """Calculate guard cell measurements.
        
        Args:
            area_wst: Whole stomata area
            area_st: Stomata pore area
            length_wst: Whole stomata length
            width_wst: Whole stomata width
            width_st: Stomata pore width
            pixel_size: Pixels per 0.1mm
            
        Returns:
            Dictionary with guard cell metrics
        """
        gc_area = area_wst - area_st
        gc_length = length_wst / (pixel_size / 100)
        gc_width = 0.5 * (width_wst - width_st) / (pixel_size / 100)
        aperture_width = width_st / (pixel_size / 100)
        
        return {
            'guardCell_area': gc_area,
            'guardCell_length': gc_length,
            'guardCell_width': gc_width,
            'aperture_width': aperture_width,
            'ratio_pore_to_gc': area_st / gc_area if gc_area > 0 else 0
        }
    
    @staticmethod
    def format_series_for_export(series: pd.Series, decimal_places: int = 0) -> pd.Series:
        """Format a pandas Series for export to Excel.
        
        Args:
            series: Input series
            decimal_places: Number of decimal places
            
        Returns:
            Formatted series
        """
        if decimal_places == 0:
            return series.map('{:,.0f}'.format)
        else:
            return series.map(f'{{:,.{decimal_places}f}}'.format)
