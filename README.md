Metadata Removal Tool
Description
This Python script provides a straightforward solution for removing metadata from various file types including images (JPEG, PNG, etc.), PDFs, DOCX files, MP3 and FLAC audio files, XLSX spreadsheets, and ZIP archives. Utilizing a simple graphical user interface (GUI) built with Tkinter, users can select files from which they wish to strip metadata, enhancing privacy and security.

Warning
Important: This tool overwrites the original files with their metadata-stripped versions. It is strongly recommended to copy the files you wish to process into a temporary directory before using this tool. Future updates may include an option to avoid this behavior, but for now, please proceed with caution to avoid unintended loss of data.

Installation
This script requires Python3 and several third-party libraries. To install the required libraries, run the following command:
pip install Pillow PyPDF2 python-docx mutagen openpyxl zipfile36
Tested with 3.11.8 you can check yours with python --version

Usage
To use the Metadata Removal Tool, follow these steps:

Ensure all dependencies are installed as described in the Installation section.
Launch the script by navigating to the directory containing the script and running:
python metadata_removal_tool.py

1. Click the "Select Files" button in the GUI that appears and select the files from which you want to remove metadata.
2. The tool will process the selected files and display a success or error message for each file processed

Supported File Types
Images: .jpg, .jpeg, .png, .gif, .bmp, .tiff
PDFs: .pdf
Word Documents: .docx
Audio Files: .mp3, .flac
Excel Spreadsheets: .xlsx
ZIP Archives: .zip
Contributing
Feedback and contributions to this project are welcome. Please feel free to submit issues or pull requests with improvements or suggestions.

License
MIT License - Feel free to use, modify, and distribute as you see fit.
