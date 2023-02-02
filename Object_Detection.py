import cv2
import numpy as np


net = cv2.dnn.readNet('yolov3_training_last.weights', 'yolov3_testing.cfg')
