# filepath: setup.py
"""
Excel Diff Application Setup
Build: python setup.py build
"""

from setuptools import setup, find_packages
import os

# README 파일이 있으면 내용 읽기
readme_file = os.path.join(os.path.dirname(__file__), "README.md")
long_description = ""
if os.path.exists(readme_file):
    with open(readme_file, "r", encoding="utf-8") as f:
        long_description = f.read()

setup(
    name="excel-diff-app",
    version="1.0.0",
    author="Excel Diff Team",
    description="Excel 파일 비교 도구 - 변경된 셀을 색상으로 표시",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "openpyxl>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pyinstaller>=5.0",
        ]
    },
    entry_points={
        "console_scripts": [
            "excel-diff=excel_diff_app.__main__:main",
        ],
        "gui_scripts": [
            "excel-diff-gui=excel_diff_app.ui.main_window:run_gui",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
)