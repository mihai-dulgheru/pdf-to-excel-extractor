import os
import shutil

import PyInstaller.__main__

PyInstaller.__main__.run([
    "main.py",
    "--noconfirm",
    "--log-level=WARN",
    "--clean",
    "--name=PDFToExcelApp",
    "--add-data=assets;assets",
    "--add-data=config;config",
    "--add-data=functions;functions",
    "--add-data=modules;modules",
    "--add-data=README.md;.",
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    "--hidden-import=pdfplumber",
    "--hidden-import=PyQt6",
    "--hidden-import=requests",
    "--hidden-import=bs4",
    "--exclude-module=PyQt6.Qt3DCore",
    "--exclude-module=PyQt6.Qt3DAnimation",
    "--exclude-module=PyQt6.Qt3DExtras",
    "--exclude-module=PyQt6.Qt3DInput",
    "--exclude-module=PyQt6.Qt3DLogic",
    "--exclude-module=PyQt6.Qt3DRender",
    "--exclude-module=PyQt6.QtWebEngineCore",
    "--exclude-module=PyQt6.QtWebEngineWidgets",
    "--exclude-module=PyQt6.QtWebEngineQuick",
    "--windowed",
    "--noupx",
    "--icon=assets/icon.ico",
])

dist_dir = os.path.join(os.getcwd(), "dist", "PDFToExcelApp")
shutil.copytree("assets", os.path.join(dist_dir, "assets"), dirs_exist_ok=True)
