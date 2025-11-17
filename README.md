# WIAScanGUI

**This graphical program helps legacy flatbed scanners be supported on Windows 11 and later. The program utilizes Windows Image Acquisition (WIA) to interface with scanners. This program is primarily intended to scan photos to a local directory without needing to crop, rotate, adjust EXIF data, or color correct the images afterward. The program allows you to do all of this, with some automated features.**

### Features

- Automatic cropping of multiple images with option to manually crop.
- Automatic white balance after user calibration
- Set DPI and scanning distance
- Set timestamp and title for each scanned image.
- Rotate images using buttons

### Installation

1. Set up and activate  a python virtual environment (venv) outside of cloud storage, e.g. `C:\Users\user\venvs\EpsonScannerSoftware\.venv`.

2. Remember to install all pip package requirements using `pip install -r requirements.txt`.

3. Download [ExifTool](https://exiftool.org) for Windows. Make a directory called `tools` in the root program directory. In the `tools` directory, place `exiftool.exe` and `exiftool_files` (directory).

4. Run the Python program or use the following command to build the .exe file:
`pyinstaller WIAScanGUI.spec`.

