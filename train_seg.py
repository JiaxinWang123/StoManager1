if __name__ == '__main__':
    from ultralytics import YOLO
  

    import torch
    torch.cuda.empty_cache()


    import os
    os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"
    # Load a model
    # model = YOLO('yolov8n-seg.yaml')  # build a new model from YAML
    model = YOLO('yolov8x-seg.pt')  # load a pretrained model (recommended for training)
    # model = YOLO('yolov8n-seg.yaml')._load('yolov8n.pt')  # build from YAML and transfer weights

    # Train the model
    results3 = model.train(data='D:/YOLOV8/leaf_stomata.v8i.yolov8/data.yaml', epochs=5, imgsz=640, batch=2, fliplr=0.0, device="0",workers=0, amp =True)
