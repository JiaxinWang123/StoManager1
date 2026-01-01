from setuptools import setup, find_packages

setup(
    name="StoManager_Project",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "PyQt5",
        "ultralytics",
        "shapely",
        "opencv-python",
        "pandas",
        "scipy",
        "qtpy",
        "torch",
        "numpy",
    ],
    entry_points={
        "console_scripts": [
            "stomanager=main:main",
            "stomanager-cli=cli:main",
        ],
    },
    author="Jiaxin Wang",
    description="A tool for stomatal detection and measurement using YOLOv8-seg-x.",
    python_requires=">=3.8",
)
