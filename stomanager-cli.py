#!/usr/bin/env python3
"""
StoManager1 - Command Line Interface

A CLI tool for stomata detection and analysis without GUI dependencies.

Author: Jiaxin Wang
Contact: jiaxinwang362@gmail.com; jiaxinwang@cornell.edu
"""

import argparse
import sys
import os

# Add parent directory to path to allow imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.constants import (
    DEFAULT_PIXEL_SIZE,
    DEFAULT_CONFIDENCE,
    DEFAULT_EPOCHS,
    DEFAULT_IMAGE_SIZE,
    DEFAULT_BATCH_SIZE,
    DEFAULT_FLIPLR,
    DEFAULT_WORKERS
)
from processors.yolov3_processor import YOLOv3Processor
from processors.yolov8_processor import YOLOv8Processor
from processors.file_normalizer import FileNormalizer
from stats.calculator import StatisticsCalculator
from utils.file_utils import FileManager


class CLIProgressCallback:
    """Simple progress callback for CLI."""
    
    def __init__(self, total=100):
        self.total = total
        self.last_percent = -1
    
    def __call__(self, percent):
        """Update progress display."""
        if percent != self.last_percent:
            self.last_percent = percent
            bar_length = 50
            filled = int(bar_length * percent / 100)
            bar = '█' * filled + '-' * (bar_length - filled)
            print(f'\rProgress: |{bar}| {percent}%', end='', flush=True)
            if percent >= 100:
                print()  # New line when complete


def process_images(args):
    """Process images with YOLOv8 or YOLOv3."""
    print(f"\n{'='*60}")
    print("StoManager1 - Image Processing")
    print(f"{'='*60}")
    
    # Validate inputs
    file_manager = FileManager()
    
    if not file_manager.has_image_files(args.input):
        print(f"Error: No image files found in {args.input}")
        return 1
    
    if file_manager.has_subfolders(args.input):
        print(f"Error: Input folder contains subfolders. Please remove them.")
        return 1
    
    # Create output directory
    file_manager.create_directory(args.output)
    
    print(f"Input folder: {args.input}")
    print(f"Output folder: {args.output}")
    print(f"Model: {'YOLOv8-seg-x' if args.model == 'yolov8' else 'YOLOv3'}")
    print(f"Pixel size: {args.pixel_size}")
    print(f"Confidence: {args.confidence}")
    print()
    
    # Process
    progress = CLIProgressCallback()
    
    try:
        if args.model == 'yolov8':
            processor = YOLOv8Processor()
            processor.process_folder(
                args.input,
                args.output,
                args.pixel_size,
                args.confidence,
                progress
            )
        else:
            processor = YOLOv3Processor()
            processor.process_folder(
                args.input,
                args.output,
                args.pixel_size,
                args.confidence,
                progress
            )
        
        print("\n✓ Processing complete!")
        print(f"Results saved to: {args.output}")
        return 0
        
    except Exception as e:
        print(f"\n✗ Error during processing: {e}")
        return 1


def calculate_statistics(args):
    """Calculate statistics from processed results."""
    print(f"\n{'='*60}")
    print("StoManager1 - Statistical Analysis")
    print(f"{'='*60}")
    
    print(f"Output folder: {args.output}")
    print(f"Model: {'YOLOv8-seg-x' if args.model == 'yolov8' else 'YOLOv3'}")
    print(f"Group analysis: {args.group}")
    print()
    
    progress = CLIProgressCallback()
    calculator = StatisticsCalculator()
    
    try:
        if args.model == 'yolov8':
            if args.group:
                calculator.calculate_group_statistics_yolov8(
                    args.output,
                    progress
                )
            else:
                calculator.calculate_statistics_yolov8(
                    args.output,
                    progress
                )
        else:
            if args.group:
                print("Note: Group analysis for YOLOv3 is only available for Populus dataset")
            calculator.calculate_statistics_yolov3(
                args.output,
                progress
            )
        
        print("\n✓ Statistical analysis complete!")
        
        # Find output file
        file_manager = FileManager()
        if args.model == 'yolov8':
            csv_path = file_manager.get_output_csv_path(args.output, "yolov8")
        else:
            csv_path = args.output
        
        xlsx_files = file_manager.get_xlsx_files(csv_path)
        if xlsx_files:
            print(f"Statistics saved to: {xlsx_files[-1]}")
        
        return 0
        
    except Exception as e:
        print(f"\n✗ Error during statistical analysis: {e}")
        return 1


def normalize_files(args):
    """Normalize filenames in input folder."""
    print(f"\n{'='*60}")
    print("StoManager1 - File Normalization")
    print(f"{'='*60}")
    
    print("Note: This function is currently designed for Populus dataset")
    print(f"Input folder: {args.input}")
    print()
    
    file_manager = FileManager()
    if not file_manager.has_image_files(args.input):
        print(f"Error: No image files found in {args.input}")
        return 1
    
    try:
        normalizer = FileNormalizer()
        renamed, skipped = normalizer.normalize_folder(args.input)
        
        print(f"✓ Files renamed: {renamed}")
        print(f"✓ Files skipped: {skipped}")
        return 0
        
    except Exception as e:
        print(f"\n✗ Error during normalization: {e}")
        return 1


def train_model(args):
    """Train YOLOv8 model."""
    print(f"\n{'='*60}")
    print("StoManager1 - Model Training")
    print(f"{'='*60}")
    
    from ultralytics import YOLO
    
    print(f"Data file: {args.data}")
    print(f"Epochs: {args.epochs}")
    print(f"Image size: {args.imgsz}")
    print(f"Batch size: {args.batch}")
    print(f"Workers: {args.workers}")
    print(f"Device: {args.device}")
    print(f"AMP: {args.amp}")
    print()
    
    # Check if data file exists
    if not os.path.exists(args.data):
        print(f"Error: Data file not found: {args.data}")
        return 1
    
    try:
        # Load model
        weights = args.weights if args.weights else 'yolov8x-seg.pt'
        model = YOLO(weights)
        
        print(f"Training with {weights}...")
        
        # Train
        results = model.train(
            data=args.data,
            epochs=args.epochs,
            imgsz=args.imgsz,
            batch=args.batch,
            fliplr=args.fliplr,
            workers=args.workers,
            device=args.device,
            amp=args.amp,
            project=args.project if args.project else 'runs/segment',
            name=args.name if args.name else 'train',
            exist_ok=True
        )
        
        print("\n✓ Training complete!")
        print(f"Results saved to: {results.save_dir}")
        return 0
        
    except Exception as e:
        print(f"\n✗ Error during training: {e}")
        return 1


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='StoManager1 - Stomata Detection and Analysis Tool (CLI)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process images with YOLOv8
  %(prog)s process -i ./input -o ./output -m yolov8
  
  # Process with custom parameters
  %(prog)s process -i ./input -o ./output -p 500 -c 0.3
  
  # Calculate statistics
  %(prog)s stats -o ./output -m yolov8
  
  # Calculate grouped statistics
  %(prog)s stats -o ./output -m yolov8 --group
  
  # Normalize filenames
  %(prog)s normalize -i ./input
  
  # Train model
  %(prog)s train -d data.yaml --epochs 100 --batch 4

For questions or issues: jiaxinwang362@gmail.com
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Process command
    process_parser = subparsers.add_parser('process', help='Process images for stomata detection')
    process_parser.add_argument('-i', '--input', required=True, help='Input folder containing images')
    process_parser.add_argument('-o', '--output', required=True, help='Output folder for results')
    process_parser.add_argument('-m', '--model', choices=['yolov8', 'yolov3'], 
                               default='yolov8', help='Model to use (default: yolov8)')
    process_parser.add_argument('-p', '--pixel-size', type=float, default=DEFAULT_PIXEL_SIZE,
                               help=f'Pixels in 0.1mm (default: {DEFAULT_PIXEL_SIZE})')
    process_parser.add_argument('-c', '--confidence', type=float, default=DEFAULT_CONFIDENCE,
                               help=f'Confidence threshold (default: {DEFAULT_CONFIDENCE})')
    
    # Statistics command
    stats_parser = subparsers.add_parser('stats', help='Calculate statistics from results')
    stats_parser.add_argument('-o', '--output', required=True, help='Output folder with processed results')
    stats_parser.add_argument('-m', '--model', choices=['yolov8', 'yolov3'],
                             default='yolov8', help='Model used for processing (default: yolov8)')
    stats_parser.add_argument('-g', '--group', action='store_true',
                             help='Perform grouped analysis (Populus dataset only)')
    
    # Normalize command
    normalize_parser = subparsers.add_parser('normalize', help='Normalize filenames')
    normalize_parser.add_argument('-i', '--input', required=True, help='Input folder containing images')
    
    # Train command
    train_parser = subparsers.add_parser('train', help='Train YOLOv8 model')
    train_parser.add_argument('-d', '--data', required=True, help='Path to data.yaml file')
    train_parser.add_argument('-e', '--epochs', type=int, default=DEFAULT_EPOCHS,
                             help=f'Number of epochs (default: {DEFAULT_EPOCHS})')
    train_parser.add_argument('--imgsz', type=int, default=DEFAULT_IMAGE_SIZE,
                             help=f'Image size (default: {DEFAULT_IMAGE_SIZE})')
    train_parser.add_argument('-b', '--batch', type=int, default=DEFAULT_BATCH_SIZE,
                             help=f'Batch size (default: {DEFAULT_BATCH_SIZE})')
    train_parser.add_argument('--fliplr', type=float, default=DEFAULT_FLIPLR,
                             help=f'Flip left-right probability (default: {DEFAULT_FLIPLR})')
    train_parser.add_argument('--workers', type=int, default=DEFAULT_WORKERS,
                             help=f'Number of workers (default: {DEFAULT_WORKERS})')
    train_parser.add_argument('--device', default='0',
                             help='Device to use (0, 1, 2, cpu) (default: 0)')
    train_parser.add_argument('--amp', type=bool, default=True,
                             help='Use automatic mixed precision (default: True)')
    train_parser.add_argument('-w', '--weights', help='Path to pretrained weights (default: yolov8x-seg.pt)')
    train_parser.add_argument('--project', help='Project directory (default: runs/segment)')
    train_parser.add_argument('--name', help='Experiment name (default: train)')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Show help if no command
    if not args.command:
        parser.print_help()
        return 1
    
    # Route to appropriate function
    if args.command == 'process':
        return process_images(args)
    elif args.command == 'stats':
        return calculate_statistics(args)
    elif args.command == 'normalize':
        return normalize_files(args)
    elif args.command == 'train':
        return train_model(args)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
