# StoManager Project

This project is a reorganized version of the StoManager application, designed for better maintainability and modularity.

## Project Structure

- `main.py`: The entry point of the application.
- `src/`: Contains the source code.
  - `ui/`: Contains the UI-related classes and components.
    - `main_window.py`: The main application window.
    - `training_window.py`: The training window.
    - `common_imports.py`: Shared imports for UI components.
  - `utils/`: Contains utility functions.
    - `file_ops.py`: File and directory operations.
- `assets/`: Contains static assets like icons and images.
- `data/`: Directory for storing application data.

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

## Fixes Applied

- Fixed `SyntaxWarning: invalid escape sequence '\s'` by using raw strings for regex.
- Fixed `NameError: name 'StoManager1' is not defined` by properly passing the window instance and reorganizing the class structure.
- Modularized the code into separate files for better maintainability.
