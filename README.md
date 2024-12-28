

# Metadata Removal Tool

## Description
This Python script is like a vacuum for your files, stealthily removing metadata from a variety of file types. Whether it's images, PDFs, DOCX files, audio files, PowerPoint presentations, OpenDocument files, EPUBs, RTF files, or ZIP archives, this tool has you covered. With a simple GUI built on Tkinter and enhanced with concurrency for better performance, cleaning your files for anonymity is really easy (though processing large batches might take a little time).

## Warning
**Important**: This tool is like a magic eraser—it permanently removes metadata and overwrites the original files. To avoid any "Oops!" moments, consider duplicating your files into a temporary folder before unleashing this tool. I am planning to add a mode that doesn't overwrite files in the future, but for now, proceed with caution (also laziness 😉).

## Installation
Before you start, make sure you have Python 3 installed. This script uses several third-party libraries, so let's get them ready with this spell:

```bash
pip install Pillow PyPDF2 python-docx mutagen openpyxl piexif python-pptx odfpy
```

Double-check your Python version with `python --version` to ensure compatibility. (Tested on 3.11.8)

**Note**: 
- To handle legacy PowerPoint files (.ppt), you may need additional tools like `unoconv` or `LibreOffice` in headless mode, as they are not directly supported by the current script.
- Logging is implemented and outputs to `metadata_removal.log` in the script's directory for better traceability.

## Usage
To start your metadata-removal tool, follow these steps:

1. **Install Dependencies**: Ensure all dependencies are installed as mentioned above.
2. **Run the Script**: Navigate to the script's directory in your terminal and cast the following spell:

    ```bash
    python metadata_removal_tool.py
    ```

3. **Select Files**: A mystical window will appear. Command it by clicking "Select Files" and choosing the files you wish to cleanse.
4. **Watch the Magic**: Observe as the tool works its magic, notifying you of its victories and defeats with each file processed.

## Supported File Types
This tool handles:

- **Images**: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`
- **PDFs**: `.pdf`
- **Word Documents**: `.docx`
- **PowerPoint Presentations**: `.pptx` *(requires `python-pptx`)*
- **OpenDocument Files**: `.odt`, `.ods` *(requires `odfpy`)*
- **Audio Files**: `.mp3`, `.flac`
- **Excel Spreadsheets**: `.xlsx`
- **EPUBs**: `.epub`
- **RTF Files**: `.rtf`
- **ZIP Archives**: `.zip`

## Contributing
Feel free to request pull requests and leave your issues or improvements. I will happily help!

## License
[GNU General Public License](LICENSE) - Feel free to review the license terms in the linked file. Thank you for your interest in my project!

---

### **Key Updates & Enhancements**

1. **Expanded File Support**:
   - **PowerPoint Presentations**: Added support for `.pptx` files using `python-pptx`.
   - **OpenDocument Files**: Added support for `.odt` and `.ods` files using `odfpy`.
   - **EPUBs & RTFs**: Included support for `.epub` and `.rtf` files.
   - **ZIP Archives**: Enhanced handling of `.zip` files to remove metadata from contained files.

2. **Robust JPEG Metadata Removal**:
   - **Two-Step Approach**:
     - **Step 1**: Attempts to remove EXIF metadata using `piexif.remove()`, preserving image quality by avoiding re-encoding.
     - **Step 2**: Verifies if any EXIF data remains. If so, it falls back to re-encoding the image with Pillow to ensure all metadata is stripped, albeit with a slight risk of quality alteration.

3. **Concurrency for Enhanced Performance**:
   - Utilizes `ThreadPoolExecutor` from `concurrent.futures` to process multiple files in parallel, significantly improving performance for large batches.

4. **Comprehensive Logging**:
   - Implements Python’s `logging` module to log informational messages, warnings, and exceptions both to the console and a log file (`metadata_removal.log`), aiding in easier debugging and maintenance.

5. **User-Friendly GUI Enhancements**:
   - **Menu Bar**: Includes "Open" and "Exit" options for standard navigation.
   - **Listbox**: Displays all selected files, providing clarity on what’s being processed.
   - **Progress Bar**: Visually represents the processing progress.
   - **Status Label**: Updates users on the current state, such as the number of selected files and processing completion.

6. **Error Handling & User Feedback**:
   - Provides immediate feedback through Tkinter’s message boxes for successes, partial successes, and errors.
   - Logs detailed error messages to `metadata_removal.log` for traceability.

7. **Cutified Comments**:
   - Added fun emoticons and clear explanations within the code to make it more engaging while maintaining professionalism.

### **Final Notes**
- **Testing**: Before deploying, thoroughly test the script with various file types to ensure metadata removal works as expected.
- **Backup**: Always keep backups of original files before bulk processing, especially since re-encoding (even with minimal quality loss) cannot be undone.
- **Extensibility**: The script is structured to allow easy addition of more file formats or more sophisticated metadata removal techniques as needed.

Enjoy your fully enhanced, production-ready metadata removal tool! 🎉

