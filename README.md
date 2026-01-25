# StoManager Project

Complete reorganization of the StoManager1 stomata detection and analysis tool.

## 📁 Project Structure

```
StoManager1/
├── main.py                          # GUI application entry point
├── cli.py                           # Command-line interface (NEW!)
├── setup.py                         # Package installation
├── config/
│   ├── __init__.py
│   └── constants.py                 # All configuration constants
├── utils/
│   ├── __init__.py
│   ├── file_utils.py               # File operations
│   └── image_utils.py              # Image utilities
├── ui/
│   ├── __init__.py
│   ├── dialogs.py                  # Message boxes and dialogs
│   ├── main_window.py              # Main application window
│   └── training_window.py          # Model training interface
├── processors/
│   ├── __init__.py
│   ├── file_normalizer.py          # Filename normalization
│   ├── stomata_analyzer.py         # Analysis algorithms
│   ├── yolov3_processor.py         # YOLOv3 detection
│   └── yolov8_processor.py         # YOLOv8 segmentation
├── stats/
│   ├── __init__.py
│   └── calculator.py               # Statistical calculations
└── assets/
    ├── StoManager.ico
    ├── best.pt                      # YOLOv8 weights
    ├── yolov3_training_last.weights # YOLOv3 weights
    └── yolov3_testing.cfg           # YOLOv3 config
```

## 🚀 Installation

### Step 1: Clone the Repository

First, clone the repository from GitHub:

```bash
git clone https://github.com/JiaxinWang123/StoManager1.git
cd StoManager
```

### Step 2: Install Dependencies

#### GUI Version (Desktop Use)

Install with GUI dependencies:

```bash
pip install -e .[gui]
```

Or install dependencies manually:

```bash
pip install PyQt5 ultralytics shapely opencv-python pandas scipy qtpy torch numpy
```

#### CLI Version (Server/Headless Use)

Install without GUI dependencies:

```bash
pip install -e .
```

For headless servers, use opencv-python-headless:

```bash
pip install -e .
pip uninstall opencv-python -y
pip install opencv-python-headless
```

**Note:** The `-e` flag installs in "editable" mode, which allows you to modify the code and see changes immediately without reinstalling.

## 💻 How to Run

### GUI Application

Run the graphical interface:

```bash
python main.py
```

Or if installed via `setup.py`:

```bash
stomanager1
```

### Command-Line Interface (CLI)

The CLI allows you to run StoManager1 on servers or in batch processing workflows without a GUI.

#### Quick Start

```bash
# Get help
python cli.py --help

# Process images with YOLOv8
python cli.py process -i /input -o /output -m yolov8

# Calculate statistics
python cli.py stats -o /output -m yolov8
```

#### Available Commands

**1. Process Images**
```bash
# Basic processing
python cli.py process -i /input_folder -o /output_folder

# With custom parameters
python cli.py process -i /input -o /output -m yolov8 -p 465 -c 0.25

# Using YOLOv3
python cli.py process -i /input -o /output -m yolov3
```

**2. Calculate Statistics**
```bash
# Non-grouped analysis
python cli.py stats -o /output_folder -m yolov8

# Grouped analysis (for Populus dataset)
python cli.py stats -o /output_folder -m yolov8 --group
```

**3. Normalize Filenames**
```bash
# Normalize filenames (Populus dataset)
python cli.py normalize -i /input_folder
```

**4. Train Model**
```bash
# Train YOLOv8 model
python cli.py train -d data.yaml --epochs 100 --batch 4
```

#### CLI Parameters

**Process Command:**
- `-i, --input`: Input folder containing images (required)
- `-o, --output`: Output folder for results (required)
- `-m, --model`: Model to use: `yolov8` or `yolov3` (default: yolov8)
- `-p, --pixel-size`: Pixels in 0.1mm (default: 465)
- `-c, --confidence`: Confidence threshold (default: 0.25)

**Stats Command:**
- `-o, --output`: Output folder with processed results (required)
- `-m, --model`: Model used for processing (default: yolov8)
- `-g, --group`: Perform grouped analysis (Populus dataset only)

**Train Command:**
- `-d, --data`: Path to data.yaml file (required)
- `-e, --epochs`: Number of training epochs (default: 1000)
- `-b, --batch`: Batch size (default: 2)
- `--imgsz`: Image size (default: 640)
- `--device`: Device to use: 0, 1, 2, or cpu (default: 0)

### Server Deployment

#### Complete Setup from Scratch

```bash
# 1. Clone repository
git clone https://github.com/JiaxinWang123/StoManager.git
cd StoManager

# 2. Create virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# 3. Install for CLI only
pip install -e .

# 4. For headless servers, use opencv-python-headless
pip uninstall opencv-python -y
pip install opencv-python-headless

# 5. Verify installation
python cli.py --help
```

#### Example Workflow on Server

```bash
# 1. Create directories
mkdir -p ~/stomata_data/input ~/stomata_data/output

# 2. Upload images to server (from local machine)
scp -r /local/path/images/* user@server:~/stomata_data/input/

# 3. Process images
python cli.py process \
    -i ~/stomata_data/input \
    -o ~/stomata_data/output \
    -m yolov8

# 4. Calculate statistics
python cli.py stats \
    -o ~/stomata_data/output \
    -m yolov8

# 5. Download results (from local machine)
scp -r user@server:~/stomata_data/output ./results/
```

#### Running in Background

For long-running jobs:

```bash
# Using nohup
nohup python cli.py process -i ./input -o ./output > analysis.log 2>&1 &

# Using screen
screen -S stomata
python cli.py process -i ./input -o ./output
# Detach: Ctrl+A then D
# Reattach: screen -r stomata

# Using tmux
tmux new -s stomata
python cli.py process -i ./input -o ./output
# Detach: Ctrl+B then D
# Reattach: tmux attach -t stomata
```

## Downloads
- **Windows Application**: [Zenodo](https://doi.org/10.5281/zenodo.7686022) | [Figshare](http://doi.org/10.6084/m9.figshare.22205020)
- **Toy Dataset**: [Zenodo](https://zenodo.org/records/10553682/files/Toy%20dataset.rar?download=1)

## Documentation
- **Manual**: [StoManager1.0.0 Manual](https://github.com/JiaxinWang123/StoManager1/blob/main/StoManager1.0.0_Manual.pdf)
- **Preprint Manuscript**: [arXiv:2304.10450](https://arxiv.org/abs/arXiv:2304.10450)

## Citation
If you use StoManager1 in your research, please cite as follows:

- **Research Article**:
  - Jiaxin Wang*, Heidi J. Renninger, Qin Ma, Shichao Jin, Measuring stomatal and guard cell metrics for plant physiology and growth using StoManager1, Plant Physiology, (2024). https://doi.org/10.1093/plphys/kiae049

    Free access to [fulltext](https://academic.oup.com/plphys/advance-article/doi/10.1093/plphys/kiae049/7595555?utm_source=authortollfreelink&utm_campaign=plphys&utm_medium=email&guestAccessKey=2a1005fc-39d4-493a-b09e-123d8962bb1c).

- **Preprint**:
```bibtex
@misc{wang2023stomanager1,
      title={StoManager1: Automated, High-throughput Tool to Measure Leaf Stomata Using Convolutional Neural Networks}, 
      author={Jiaxin Wang and Heidi J. Renninger and Qin Ma},
      year={2023},
      eprint={2304.10450},
      archivePrefix={arXiv},
      primaryClass={q-bio.TO}
}
```

## Datasets
- **Labeled Hardwood and *Populus* Datasets**: Approximately 11,000 images with corresponding labels available on [figshare](https://doi.org/10.6084/m9.figshare.22255873) and [Zenodo](https://doi.org/10.5281/zenodo.8266240).
- **Code for Datasets**: [ScientificData_Labeled_Hardwood_Images](https://github.com/JiaxinWang123/ScientificData_Labeled_Hardwood_Images)
- **Publication Status**: Wang, J., Renninger, H.J. & Ma, Q. Labeled temperate hardwood tree stomatal image datasets from seven taxa of _Populus_ and 17 hardwood species. Sci Data 11, 1 (2024). https://doi.org/10.1038/s41597-023-02657-3.

## StoManager1_v.1.0.0: Enhanced Version
StoManager1_v.1.0.0 is an enhanced version of StoManager1, featuring additional stomatal metrics measured with geometrical algorithms.

- **Latest Standalone Windows Version Apps**: [Zenodo](https://doi.org/10.5281/zenodo.7686022) | [Figshare](http://doi.org/10.6084/m9.figshare.22205020)
  
![StoManager1_v 1 0 0](https://github.com/JiaxinWang123/StoManager1/assets/98176596/2e15a57f-1de9-409d-b888-71adc53524ed)

## Features

### GUI Application
- 🖼️ Interactive image viewing and navigation
- 📊 Real-time visualization of detection results
- 🎯 Adjustable detection parameters
- 📈 Statistical analysis with visual feedback
- 🔧 Model training interface
- 💾 Export results to CSV and Excel

### Command-Line Interface (NEW!)
- 🚀 Server and cluster deployment
- 📦 Batch processing of large datasets
- 🔄 Automated workflows and pipelines
- 💻 Headless operation (no GUI required)
- ⚡ High-performance processing
- 📝 Progress tracking and logging

## Troubleshooting

### CLI Command Not Found

If `stomanager1-cli` command is not recognized:

**Solution 1: Run directly with Python**
```bash
python cli.py --help
```

**Solution 2: Add to PATH**
```bash
# Linux/Mac
export PATH="$PATH:$HOME/.local/bin"

# Windows PowerShell
$env:Path += ";C:\Users\YourName\AppData\Local\Programs\Python\Python3X\Scripts"
```

### Module Import Errors

Ensure all `__init__.py` files exist:
```bash
touch config/__init__.py
touch processors/__init__.py
touch stats/__init__.py
touch utils/__init__.py
```

### GPU/CUDA Issues

Check CUDA availability:
```bash
python -c "import torch; print('CUDA available:', torch.cuda.is_available())"
```

For headless servers, use CPU:
```bash
# The CLI will automatically use GPU if available, otherwise CPU
# To force CPU, you can modify the device parameter in processor code
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the terms specified in the repository.

## Contact

- **Author**: Jiaxin Wang
- **Email**: jiaxinwang362@gmail.com; jiaxinwang@cornell.edu
- **GitHub**: [https://github.com/JiaxinWang123/StoManager1](https://github.com/JiaxinWang123/StoManager1)
