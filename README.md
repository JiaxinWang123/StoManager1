# StoManager Project

Complete reorganization of the StoManager1 stomata detection and analysis tool.

## 📁 Project Structure

```
StoManager1/
├── main.py                          # Application entry point
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
│   ├── main_window.py              # Main application window (2 parts)
│   └── training_window.py          # Model training interface
├── processors/
│   ├── __init__.py
│   ├── file_normalizer.py          # Filename normalization
│   ├── stomata_analyzer.py         # Analysis algorithms
│   ├── yolov3_processor.py         # YOLOv3 detection
│   └── yolov8_processor.py         # YOLOv8 segmentation
├── statistics/
│   ├── __init__.py
│   └── calculator.py               # Statistical calculations
└── resources/
    ├── StoManager.ico
    ├── best.pt                      # YOLOv8 weights
    ├── yolov3_training_last.weights # YOLOv3 weights
    └── yolov3_testing.cfg           # YOLOv3 config
```

## 🚀 Installation

### Step 1: Create Directory Structure

```bash
mkdir -p StoManager1/{config,utils,ui,processors,statistics,resources}
```

### Step 2: Create __init__.py Files

```bash
# Create empty __init__.py in each package
touch StoManager1/config/__init__.py
touch StoManager1/utils/__init__.py
touch StoManager1/ui/__init__.py
touch StoManager1/processors/__init__.py
touch StoManager1/statistics/__init__.py
```

### Step 3: Copy Files

Copy each file from the artifacts into its respective location:

1. **config/constants.py** - Configuration constants
2. **utils/file_utils.py** - File management utilities
3. **utils/image_utils.py** - Image processing utilities
4. **ui/dialogs.py** - Dialog management
5. **ui/training_window.py** - Training interface
6. **ui/main_window.py** - Combine Part 1 and Part 2 into single file
7. **processors/file_normalizer.py** - Filename normalization
8. **processors/stomata_analyzer.py** - Analysis algorithms
9. **processors/yolov3_processor.py** - YOLOv3 processing
10. **processors/yolov8_processor.py** - YOLOv8 processing
11. **statistics/calculator.py** - Statistics calculator
12. **main.py** - Application entry point

## Installation

You can install the project and its dependencies using `setup.py`:

```bash
pip install -e .
```

Alternatively, you can install the dependencies manually:

```bash
pip install PyQt5 ultralytics shapely opencv-python pandas scipy qtpy torch numpy
```

## How to Run

Run the application using the entry point:

```bash
python main.py
```

Or if installed via `setup.py`:

```bash
stomanager
```

## CLI Version (Server/No-UI)

You can run the tool without a UI using the CLI version:

```bash
# For prediction
python cli.py predict --input /path/to/images --output /path/to/output --conf 0.25

# For training
python cli.py train --data /path/to/data.yaml --epochs 100 --batch 4 --device cpu
```

If installed via `setup.py`:

```bash
stomanager-cli predict --input /path/to/images --output /path/to/output
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
