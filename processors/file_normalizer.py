"""Filename normalization for Populus dataset."""

import os
import re
from typing import List, Tuple, Optional
from config.constants import SITES, BLOCKS, YEARS, MONTHS, TAILS, CLONES


class FileNormalizer:
    """Normalizes filenames for Populus dataset."""
    
    def __init__(self):
        self.sites = SITES
        self.blocks = BLOCKS
        self.years = YEARS
        self.months = MONTHS
        self.tails = TAILS
        self.clones = CLONES
    
    def normalize_filename(self, filename: str) -> Optional[str]:
        """Normalize a single filename.
        
        Args:
            filename: Original filename
            
        Returns:
            Normalized filename or None if cannot normalize
        """
        # Remove extension
        name_without_ext = os.path.splitext(filename)[0]
        
        # Split by comma and clean whitespace
        parts = name_without_ext.split(",")
        split_name = []
        for part in parts:
            part = re.sub(r"\s+", ",", part.strip())
            split_name.append(part)
        
        split_name = ",".join(split_name).split(",")
        
        # Need at least 5 parts for normalization
        if len(split_name) < 5:
            return None
        
        # Extract components
        site = self._find_in_list(split_name, self.sites)
        block = self._find_in_list(split_name, self.blocks)
        year = self._find_in_list(split_name, self.years)
        month = self._find_in_list(split_name, self.months)
        tail = self._find_in_list(split_name, self.tails)
        clone = self._find_in_list(split_name, self.clones)
        
        # Build normalized name
        if all([site, block, clone, month, year, tail]):
            return f"{site},{block},{clone},{month},{year},{tail}.jpg"
        
        return None
    
    def _find_in_list(self, parts: List[str], target_list: List[str]) -> Optional[str]:
        """Find first matching item from target list in parts.
        
        Args:
            parts: List of filename parts
            target_list: List of valid values to match
            
        Returns:
            First matching value or None
        """
        for part in parts:
            if part in target_list:
                return part
        return None
    
    def normalize_folder(self, folder_path: str) -> Tuple[int, int]:
        """Normalize all image files in a folder.
        
        Args:
            folder_path: Path to folder containing images
            
        Returns:
            Tuple of (files_renamed, files_skipped)
        """
        files_renamed = 0
        files_skipped = 0
        
        for filename in os.listdir(folder_path):
            if not filename.lower().endswith(('.jpg', '.png', '.tif', '.jpeg')):
                continue
            
            file_path = os.path.join(folder_path, filename)
            new_filename = self.normalize_filename(filename)
            
            if new_filename:
                new_path = os.path.join(folder_path, new_filename)
                if not os.path.exists(new_path):
                    os.rename(file_path, new_path)
                    files_renamed += 1
                else:
                    files_skipped += 1
            else:
                files_skipped += 1
        
        return files_renamed, files_skipped
