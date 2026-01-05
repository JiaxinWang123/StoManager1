"""Configuration constants for StoManager1."""

# Default values
DEFAULT_PIXEL_SIZE = 465
DEFAULT_CONFIDENCE = 0.25
DEFAULT_EPOCHS = 1000
DEFAULT_IMAGE_SIZE = 640
DEFAULT_BATCH_SIZE = 2
DEFAULT_FLIPLR = 0
DEFAULT_WORKERS = 0

# File extensions
ALLOWED_IMAGE_EXTENSIONS = ('jpg', 'png', 'tif', 'jpeg')

# Model paths
YOLOV3_WEIGHTS = "assets/yolov3_training_last.weights"
YOLOV3_CONFIG = "assets/yolov3_testing.cfg"
YOLOV8_WEIGHTS = "assets/best.pt"

# Classes
STOMATA_CLASSES = ["whole_stomata", "stomata"]

# Outlier detection percentiles
OUTLIER_LOWER_PERCENTILE = 2.5
OUTLIER_UPPER_PERCENTILE = 97.5
OUTLIER_IQR_MULTIPLIER = 1.5

# UI Constants
ICON_PATH = "assets/StoManager.ico"
WINDOW_TITLE = "StoManager1 v.1.0.0."
WINDOW_WIDTH = 1280
WINDOW_HEIGHT = 720

# Site, Block, Year, Month, Tail, Clone lists for normalization
SITES = ["M", "Mon", "m", "mon", "P", "p", "Pon", "pon", "PON", 
         "M ", "Mon ", "m ", "mon ", "P ", "p ", "Pon ", "pon ", "PON "]

BLOCKS = ["b1", "b2", "b5", "B1", "B2", "B5", 
          "b1 ", "b2 ", "b5 ", "B1 ", "B2 ", "B5 ",
          " b1", " b2", " b5", " B1", " B2", " B5"]

YEARS = ["20", "21", "22", "23", "24", 
         " 20", " 21", " 22", " 23", " 24",
         "20 ", "21 ", "22 ", "23 ", "24 ",
         "2022", "2020", "2021", "2019"]

MONTHS = ["June", "june", "July", "july", "Aug", "aug", "AUG", "JULY", "JUNE",
          " June", " june", " July", " july", " Aug", " aug", " AUG", " JULY", " JUNE",
          "June ", "june ", "July ", "july ", "Aug ", "aug ", "AUG ", "JULY ", "JUNE ",
          "Sept", "Sep", " Sept", "sep", " sep"]

TAILS = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
         "Aci", "Aci1", "Aci2", "Aci3", "Aci4", "Aci5", "Aci6", "Aci7", "Aci8", "Aci9", "Aci10",
         "aci", "aci1", "aci2", "aci3", "aci4", "aci5", "aci6", "aci7", "aci8", "aci9", "aci10",
         "ACi", "aCI", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", " 9", " 10",
         "40X", "20X", " 40X", " 20x", " 40X ", " 20X ", "40x ", "20X ",
         "EC", "DN", "DM", "DD"]

CLONES = ["S7C2", "s7c2", "s7c4", "S7C4", "ST-66", "st-66", "st66", "ST66",
          "19", "22", "110412", "11412", "120-4", "3-1", "6-1", "6-5",
          "ST-70", "ST-75", "st-70", "st-75", "st70", "st75",
          "6323", "6329", "8019", "9225", "9707", "11690", "13693", "13724",
          "24033", "24056", "24066", "24114", "24120", "24159", "29310", "433"]
