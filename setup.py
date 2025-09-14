from setuptools import setup, find_packages
import os

def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname), encoding="utf-8").read()

setup(
    name="StoManager1",
    version="1.0.0",
    author="Jiaxin Wang",
    author_email="coolwjx@foxmail.com",
    description="StoManager1: Stomata detection, measurement, and YOLO-based training tool",
    long_description=read("README.md") if os.path.exists("README.md") else "",
    long_description_content_type="text/markdown",
    url="https://github.com/JiaxinWang123/StoManager1",
    py_modules=[
        "StoManager1_v10_new",
        "model_training_in_app",
        "train_seg",
        "res"
    ],
    install_requires=[
        "PyQt5>=5.15.0",
        "pandas>=1.5.0",
        "numpy<2.0",
        "opencv-python>=4.7.0",
        "Pillow>=9.0.0",
        "ultralytics>=8.0.0",
    ],
    entry_points={
        "console_scripts": [
            "stomanager=StoManager1_v10_new:main",  # main() must exist in StoManager1_v10_new.py
        ],
    },
    include_package_data=True,
    package_data={
        "": [
            "*.ico",
            "*.txt",
            "*.cfg",
            "*.qrc",
            "*.ui",
            "*.pt",
            "*.weights",
            "*.pdf"
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)
