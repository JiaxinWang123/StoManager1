import sys
from ultralytics import YOLO
import torch
import matplotlib
import matplotlib.pyplot as plt
matplotlib.use('TkAgg')
matplotlib.use('Qt5Agg')
matplotlib.use('pdf')

if __name__ == '__main__':

    if torch.cuda.is_available():
        device = "0"
    else:
        device = "cpu"

    data_input_path = sys.argv[1]
    script_path = sys.argv[2]
    epoch = int(sys.argv[3])
    imgsz = int(sys.argv[4])
    batch = int(sys.argv[5])
    fliplr = float(sys.argv[6])
    workers = int(sys.argv[7])
    weights_path = sys.argv[8]

    torch.cuda.empty_cache()
    import os
    os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

    # Load a model
    model = YOLO(weights_path)  # Load a pretrained model (recommended for training)

    # Train the model
    results = model.train(data=data_input_path, epochs=epoch, imgsz=imgsz, batch=batch, fliplr=fliplr, device=device, workers=workers)

