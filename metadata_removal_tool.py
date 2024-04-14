import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
from mutagen.mp3 import MP3
from mutagen.id3 import ID3, ID3NoHeaderError
from mutagen.flac import FLAC
from openpyxl import load_workbook
import zipfile

def remove_metadata(file_path):
    # Determine the file type (extension) ༼ つ ◕_◕ ༽つ
    file_extension = os.path.splitext(file_path)[1].lower()

    # Call the appropriate function based on the file extension (～￣▽￣)～
    if file_extension in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']:
        remove_metadata_from_image(file_path)
    elif file_extension == '.pdf':
        remove_metadata_from_pdf(file_path)
    elif file_extension == '.docx':
        remove_metadata_from_docx(file_path)
    elif file_extension == '.mp3':
        remove_metadata_from_mp3(file_path)
    elif file_extension == '.flac':
        remove_metadata_from_flac(file_path)
    elif file_extension == '.xlsx':
        remove_metadata_from_xlsx(file_path)
    elif file_extension == '.zip':
        remove_metadata_from_zip(file_path)  # Handle ZIP files ಠ╭╮ಠ
    else:
        print(f"File type {file_extension} not supported.")


def remove_metadata_from_zip(zip_path):
    try:
        # Create a temporary directory to extract the ZIP ಠಿ_ಠ
        temp_dir = "temp_zip_extract"
        os.makedirs(temp_dir, exist_ok=True)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    filePath = os.path.join(root, file)
                    zip_ref.write(filePath, os.path.relpath(filePath, temp_dir))
        
        # Clean up the temporary directory
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(temp_dir)
        
        return True
    except Exception as e:
        print(f"Error processing ZIP {zip_path}: {e}")
        return False
        

def remove_metadata_from_image(image_path):
    try:
        with Image.open(image_path) as img:
            data = img.getdata()
            clean_img = Image.new(img.mode, img.size)
            clean_img.putdata(data)
            clean_img.save(image_path)
        return True
    except Exception as e:
        print(f"Error processing image {image_path}: {e}")
        return False

def remove_metadata_from_pdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        with open(pdf_path, 'wb') as out_pdf:
            writer.write(out_pdf)
        return True
    except Exception as e:
        print(f"Error processing PDF {pdf_path}: {e}")
        return False

def remove_metadata_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        for section in doc.sections:
            section.start_type = None

        doc.core_properties.author = ""
        doc.core_properties.title = ""
        doc.core_properties.last_modified_by = ""  # Remove "Last Modified By" metadata from issue #1
        doc.save(docx_path)
        return True
    except Exception as e:
        print(f"Error processing DOCX {docx_path}: {e}")
        return False

def remove_metadata_from_mp3(mp3_path):
    try:
        audio = MP3(mp3_path, ID3=ID3)
        audio.delete()  # Remove all tags ( ﾉ ﾟｰﾟ)ﾉ
        audio.save(mp3_path)
        return True
    except ID3NoHeaderError:
        audio = MP3(mp3_path)
        audio.add_tags()
        audio.save(mp3_path)
        return True
    except Exception as e:
        print(f"Error processing MP3 {mp3_path}: {e}")
        return False

def remove_metadata_from_flac(flac_path):
    try:
        audio = FLAC(flac_path)
        audio.delete()  # Remove all tags ^_____^
        audio.save()
        return True
    except Exception as e:
        print(f"Error processing FLAC {flac_path}: {e}")
        return False

def remove_metadata_from_xlsx(xlsx_path):
    try:
        workbook = load_workbook(filename=xlsx_path)
        workbook.properties.creator = ""
        workbook.properties.title = ""
        workbook.save(xlsx_path)
        return True
    except Exception as e:
        print(f"Error processing XLSX {xlsx_path}: {e}")
        return False

def remove_metadata(file_paths):
    for file_path in file_paths:
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension in ['.jpg', '.jpeg', '.png']:
            success = remove_metadata_from_image(file_path)
        elif file_extension == '.pdf':
            success = remove_metadata_from_pdf(file_path)
        elif file_extension == '.docx':
            success = remove_metadata_from_docx(file_path)
        elif file_extension == '.mp3':
            success = remove_metadata_from_mp3(file_path)
        elif file_extension == '.xlsx':
            success = remove_metadata_from_xlsx(file_path)
        elif file_extension == '.zip':
            success = remove_metadata_from_zip(file_path)  
        else:
            messagebox.showerror("Unsupported File", f"File type {file_extension} is not supported.")
            continue

        if success:
            messagebox.showinfo("Success", f"Metadata removed from {os.path.basename(file_path)}")
        else:
            messagebox.showerror("Error", f"Failed to remove metadata from {os.path.basename(file_path)}")

def select_files():
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        remove_metadata(file_paths)

def create_ui():
    window = tk.Tk()
    window.title("Metadata Removal Tool")

    btn_select_files = tk.Button(window, text="Select Files", command=select_files)
    btn_select_files.pack(pady=20)

    window.mainloop()

if __name__ == "__main__":
    create_ui()
