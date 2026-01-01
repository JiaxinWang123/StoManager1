import argparse
import sys
from src.core import process_images, train_model

def main():
    parser = argparse.ArgumentParser(description="StoManager CLI - Stomatal Detection and Training")
    subparsers = parser.add_subparsers(dest="command", help="Commands")

    # Predict command
    predict_parser = subparsers.add_parser("predict", help="Run stomatal detection")
    predict_parser.add_argument("--input", required=True, help="Input folder containing images")
    predict_parser.add_argument("--output", required=True, help="Output folder for results")
    predict_parser.add_argument("--conf", type=float, default=0.25, help="Confidence threshold")
    predict_parser.add_argument("--p", type=int, default=465, help="Pixels in 0.1 mm")

    # Train command
    train_parser = subparsers.add_parser("train", help="Train the model")
    train_parser.add_argument("--data", required=True, help="Path to data.yaml")
    train_parser.add_argument("--epochs", type=int, default=1000, help="Number of epochs")
    train_parser.add_argument("--imgsz", type=int, default=640, help="Image size")
    train_parser.add_argument("--batch", type=int, default=2, help="Batch size")
    train_parser.add_argument("--device", default="cpu", help="Device (cpu or 0, 1, etc.)")

    args = parser.parse_args()

    if args.command == "predict":
        process_images(args.input, args.output, args.conf, args.p)
    elif args.command == "train":
        train_model(args.data, args.epochs, args.imgsz, args.batch, args.device)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
