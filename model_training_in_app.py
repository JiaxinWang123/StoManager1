import sys
from ultralytics import YOLO
import torch
import matplotlib
import matplotlib.pyplot as plt

matplotlib.use('TkAgg') # this is used to help Pyinstaller pull matplotlib backends
matplotlib.use('Qt5Agg') # this is used to help Pyinstaller pull matplotlib backends
matplotlib.use('pdf') # this is used to help Pyinstaller pull matplotlib backends

if __name__ == '__main__':

    if torch.cuda.is_available():
        device_ = "0"
    else:
        device_ = "cpu"

    data_input_path = sys.argv[1]
    script_path = sys.argv[2]
    epoch = int(sys.argv[3])
    imgsz = int(sys.argv[4])
    batch = int(sys.argv[5])
    fliplr = float(sys.argv[6])
    workers = int(sys.argv[7])
    weights_path = sys.argv[8]
    amp = eval(sys.argv[9])
    device = sys.argv[10]

    # Determine with device should be used for model training
    if device_ == "cpu" and device == "cpu":
        device = "cpu"
    elif device_ =="0" and device =="cpu":
        device = "cpu"
    elif device_ =="0" and device =="gpu":
        device = "0"
    else:
        device ="0"

    torch.cuda.empty_cache() # Clear GPU cache 
    import os
    os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

    # Load a model
    model = YOLO(weights_path)  # Load a pretrained model (recommended for training)

    # Train the model
    results = model.train(data=data_input_path, epochs=epoch, imgsz=imgsz, batch=batch, fliplr=fliplr, device=device, workers=workers, amp=amp)