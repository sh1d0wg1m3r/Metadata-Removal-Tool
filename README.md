
# Metadata Removal Tool

## Description
This Python script is like a ninja for your files, stealthily removing metadata from a variety of file types. Whether it's images, PDFs, DOCX files, or audio files, this tool has you covered. With a simple GUI built on Tkinter, cleaning your files is as easy as clicking a button.

## Warning
🚨 **Important**: This tool is like a magic eraser - it permanently removes metadata and overwrites the original files. To avoid any "Oops!" moments, consider duplicating your files into a temporary folder before unleashing this ninja. We're planning to add a "stealth mode" that doesn't overwrite files in the future, but for now, proceed with caution.

## Installation
Before you start, make sure you have Python 3.11.8 as your ally. This script uses several third-party libraries, so let's get them ready with this spell:

```bash
pip install Pillow PyPDF2 python-docx mutagen openpyxl zipfile36
```

Double-check your Python version with `python --version` to ensure compatibility.

## Usage
To start your metadata-removal quest, follow these steps:

1. Make sure you've installed all dependencies as mentioned above.
2. Run the script by navigating to its lair (directory) and casting the following spell in your terminal:

```bash
python metadata_removal_tool.py
```

3. A mystical window will appear. Command it by clicking "Select Files" and choosing the files you wish to cleanse.
4. Watch as the tool works its magic, notifying you of its victories and defeats with each file processed.

## Supported File Types
This tool is skilled in handling:
- Images: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`
- PDFs: `.pdf`
- Word Documents: `.docx`
- Audio Files: `.mp3`, `.flac`
- Excel Spreadsheets: `.xlsx`
- ZIP Archives: `.zip`

## Contributing
Got ideas to make this tool even sneakier or more powerful? Contributions are welcome! Share your spells (code improvements) and tales (feedback) through issues and pull requests.

## License
[MIT License](LICENSE) - Free as a bird! Use, modify, and distribute as your heart desires.
