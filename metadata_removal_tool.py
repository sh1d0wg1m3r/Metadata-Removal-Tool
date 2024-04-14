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

def remove_metadata_from_zip(zip_path):
    try:
        # Create a temporary directory to extract the ZIP (づ｡◕‿‿◕｡)づ
        temp_dir = "temp_zip_extract"
        os.makedirs(temp_dir, exist_ok=True)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Remove metadata from all files within the ZIP (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                remove_metadata(file_path)
        
        # Repackage the ZIP with clean files (～￣▽￣)～
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    zip_ref.write(file_path, os.path.relpath(file_path, temp_dir))
        
        # Clean up the temporary directory (ノ^_^)ノ
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
            # Convert image data to a list (✿◠‿◠)
            data = list(img.getdata())
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

        # Add pages to the writer, leaving out metadata (▀̿Ĺ̯▀̿ ̿)
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
        # Clear the metadata fields (ﾉ≧∀≦)ﾉ
        metadata_properties = [
            'author', 'comments', 'category', 'content_status',
            'identifier', 'keywords', 'language', 'last_modified_by',
            'last_printed', 'revision', 'subject', 'title', 'version'
        ]
        for prop in metadata_properties:
            try:
                setattr(doc.core_properties, prop, "")
            except ValueError:
                # Skip updating the property if it requires a specific data type
                pass
        doc.settings.odd_and_even_pages_header_footer = False
        doc.save(docx_path)
        return True
    except Exception as e:
        print(f"Error processing DOCX {docx_path}: {e}")
        return False

def remove_metadata_from_mp3(mp3_path):
    try:
        audio = MP3(mp3_path, ID3=ID3)
        audio.delete()  # Remove all tags (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
        audio.save(mp3_path)
        return True
    except ID3NoHeaderError:
        audio = MP3(mp3_path)
        audio.save(mp3_path)
        return True
    except Exception as e:
        print(f"Error processing MP3 {mp3_path}: {e}")
        return False

def remove_metadata_from_flac(flac_path):
    try:
        audio = FLAC(flac_path)
        audio.delete()  # Remove all tags (◠﹏◠)
        audio.save()
        return True
    except Exception as e:
        print(f"Error processing FLAC {flac_path}: {e}")
        return False

def remove_metadata_from_xlsx(xlsx_path):
    try:
        workbook = load_workbook(filename=xlsx_path)
        metadata_properties = [
            'creator', 'title', 'subject', 'description',
            'keywords', 'category', 'comments', 'last_modified_by',
            'company', 'manager'
        ]
        for prop in metadata_properties:
            try:
                setattr(workbook.properties, prop, "")
            except ValueError:
                pass
        workbook.save(xlsx_path)
        return True
    except Exception as e:
        print(f"Error processing XLSX {xlsx_path}: {e}")
        return False

def remove_metadata(file_path):
    # Determine the file type based on the extension (ﾉ≧∀≦)ﾉ
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']:
        return remove_metadata_from_image(file_path)
    elif file_extension == '.pdf':
        return remove_metadata_from_pdf(file_path)
    elif file_extension == '.docx':
        return remove_metadata_from_docx(file_path)
    elif file_extension == '.mp3':
        return remove_metadata_from_mp3(file_path)
    elif file_extension == '.flac':
        return remove_metadata_from_flac(file_path)
    elif file_extension == '.xlsx':
        return remove_metadata_from_xlsx(file_path)
    elif file_extension == '.zip':
        return remove_metadata_from_zip(file_path)
    else:
        print(f"File type {file_extension} not supported.")
        return False

def process_files(file_paths):
    success_count = 0
    for file_path in file_paths:
        if remove_metadata(file_path):
            success_count += 1
        else:
            messagebox.showerror("Error", f"Failed to remove metadata from {os.path.basename(file_path)}")
    
    if success_count == len(file_paths):
        messagebox.showinfo("Success", "Metadata removed from all selected files. (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧")
    else:
        messagebox.showinfo("Partial Success", f"Metadata removed from {success_count} out of {len(file_paths)} files. (╯°□°）╯︵ ┻━┻")

def select_files():
    # Open the file dialog to select files (◕‿◕✿)
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        process_files(file_paths)

def create_ui():
    # Create the main window (✿◠‿◠)
    window = tk.Tk()
    window.title("Metadata Removal Tool")

    # Create the "Select Files" button (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
    btn_select_files = tk.Button(window, text="Select Files", command=select_files)
    btn_select_files.pack(pady=20)

    # Start the main event loop (◡‿◡✿)
    window.mainloop()

if __name__ == "__main__":
    # Run the program (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
    create_ui()
