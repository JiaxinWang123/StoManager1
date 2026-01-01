import sys
import os
import re
import cv2
import csv
import glob
import math
import torch
import random
import string
import shutil
import io
import subprocess
import webbrowser
import numpy as np
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLineEdit, QPushButton, QCheckBox, QLabel, QFileDialog, QMessageBox
from PyQt5.QtGui import QPixmap, QFont, QIcon, QStandardItemModel
from qtpy.QtCore import QThread, Signal
from ultralytics import YOLO
from shapely.geometry import Polygon
from shapely import Point
from scipy.spatial import distance, distance_matrix
from scipy.sparse.csgraph import minimum_spanning_tree
