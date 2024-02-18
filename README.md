
# Metadata Removal Tool

## Description
This Python script is like a vacuum for your files, stealthily removing metadata from a variety of file types. Whether it's images, PDFs, DOCX files, or audio files, file format that I will add in the future this tool got you covered. With a simple GUI built on Tkinter, cleaning your for anonimity is really easy although a lil slow.

## Warning
**Important**: This tool is like a magic eraser - it permanently removes metadata and overwrites the original files. To avoid any "Oops!" moments, consider duplicating your files into a temporary folder before unleashing this tool. I am  planning to add a mode that doesn't overwrite files in the future, but for now, proceed with caution.( Also laziness )

## Installation
Before you start, make sure you have Python3. This script uses several third-party libraries, so let's get them ready with this spell:

```bash
pip install Pillow PyPDF2 python-docx mutagen openpyxl zipfile36
```

Double-check your Python version with `python --version` to ensure compatibility. ( Tested on 3.11.8 )

## Usage
To start your metadata-removal_tool, follow these steps:

1. Make sure you've installed all dependencies as mentioned above.
2. Run the script by navigating to its directory and casting the following spell in your terminal:

```bash
python metadata_removal_tool.py
```

3. A mystical window will appear. Command it by clicking "Select Files" and choosing the files you wish.
4. Watch as the tool works its magic, notifying you of its victories and defeats with each file processed.

## Supported File Types
This tool handles:
- Images: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`
- PDFs: `.pdf`
- Word Documents: `.docx`
- Audio Files: `.mp3`, `.flac`
- Excel Spreadsheets: `.xlsx`
- ZIP Archives: `.zip`

## Contributing
Feel free to request pull requests and leave your issues or improvements. I will happily help.

## License
[MIT License](LICENSE) - Free as a bird! Use, modify, and distribute as your heart desires.
