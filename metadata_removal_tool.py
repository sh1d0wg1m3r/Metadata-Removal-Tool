import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging

# ~(^-^)~ Pillow for images
from PIL import Image

# ~(^-^)~ PDFs
from PyPDF2 import PdfReader, PdfWriter

# ~(^-^)~ DOCX
from docx import Document

# ~(^-^)~ Audio
from mutagen.mp3 import MP3
from mutagen.id3 import ID3, ID3NoHeaderError
from mutagen.flac import FLAC

# ~(^-^)~ XLSX
from openpyxl import load_workbook

# ~(^-^)~ JPEG EXIF removal
import piexif

# ~(^-^)~ PPTX (PowerPoint 2007+); PPT is legacy
try:
    from pptx import Presentation
    CAN_HANDLE_PPTX = True
except ImportError:
    CAN_HANDLE_PPTX = False

# ~(^-^)~ ODT/ODS (OpenDocument), using odfpy if installed
try:
    from odf.opendocument import load as odf_load
    CAN_HANDLE_ODF = True
except ImportError:
    CAN_HANDLE_ODF = False

# ~(^-^)~ Logging for better production tracing
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("metadata_removal.log"),
        logging.StreamHandler()
    ]
)


################################################################
# ~(^-^)~ METADATA REMOVAL LOGIC FOR VARIOUS FILE FORMATS
################################################################

def remove_metadata_from_zip(zip_path):
    """ 
    Unzip -> Remove metadata from each entry -> Re-zip.
    (｡•̀ᴗ-)✧
    """
    temp_dir = "temp_zip_extract"
    try:
        logging.info(f"Processing ZIP: {zip_path}")
        os.makedirs(temp_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # ~(^-^)~ Clean each file inside temp_dir
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                remove_metadata(file_path)

        # ~(^-^)~ Repackage as a new, cleaned ZIP
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname=arc_name)

        # ~(^-^)~ Cleanup
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(temp_dir)

        return True
    except Exception as e:
        logging.exception(f"Error processing ZIP {zip_path}: {e}")
        return False


def remove_exif_jpeg(file_path):
    """ 
    Strip EXIF from JPEG without re-encoding 
    ヾ(⌐■_■)ノ♪
    """
    try:
        logging.info(f"Stripping EXIF from JPEG: {file_path}")
        piexif.remove(file_path)
        # Verify removal
        exif_dict = piexif.load(file_path)
        if any(exif_dict[tag] for tag in ["0th", "Exif", "GPS", "1st"]):
            logging.info(f"EXIF still present after piexif.remove. Re-encoding: {file_path}")
            _reencode_jpeg(file_path)
        else:
            logging.info(f"EXIF successfully removed: {file_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing JPEG EXIF {file_path}: {e}")
        # Fallback to re-encode if piexif fails
        try:
            _reencode_jpeg(file_path)
            return True
        except Exception as ex:
            logging.exception(f"Failed to re-encode JPEG: {file_path}, error: {ex}")
            return False


def _reencode_jpeg(file_path):
    """
    Re-encode the JPEG with Pillow. This definitely strips EXIF,
    but may alter quality/size if not configured carefully.
    """
    try:
        with Image.open(file_path) as img:
            # Convert to RGB (some images might be in different modes)
            rgb_img = img.convert("RGB")
            # Save, forcing 'exif' to be blank
            rgb_img.save(file_path, format='JPEG', exif=b'')
        logging.info(f"Re-encoded JPEG to remove residual EXIF: {file_path}")
    except Exception as e:
        logging.exception(f"Failed to re-encode JPEG: {file_path}, error: {e}")
        raise


def remove_metadata_from_jpeg(file_path):
    """
    Attempt to remove EXIF from JPEG using piexif (no re-encoding).
    If EXIF still remains, fallback to re-encoding with Pillow to ensure
    all metadata is stripped. This may alter quality/file size slightly.
    """
    try:
        remove_exif_jpeg(file_path)
        return True
    except Exception as e:
        logging.error(f"Failed to remove metadata from JPEG {file_path}: {e}")
        return False


def remove_metadata_from_image(image_path):
    """
    Images beyond JPEG (PNG, GIF, BMP, TIFF):
    Re-encode with Pillow to drop metadata.
    (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
    """
    ext = os.path.splitext(image_path)[1].lower()
    if ext in [".jpg", ".jpeg"]:
        return remove_metadata_from_jpeg(image_path)
    try:
        logging.info(f"Processing image: {image_path}")
        with Image.open(image_path) as img:
            data = list(img.getdata())
            clean_img = Image.new(img.mode, img.size)
            clean_img.putdata(data)
            # Keep the original format if possible
            clean_img.save(image_path, format=img.format)
        logging.info(f"Metadata removed from image: {image_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing image {image_path}: {e}")
        return False


def remove_metadata_from_pdf(pdf_path):
    """
    PDF: Read pages, rewrite them, omit doc info. 
    (／・ω・)／
    """
    try:
        logging.info(f"Processing PDF: {pdf_path}")
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(pdf_path, 'wb') as out_pdf:
            writer.write(out_pdf)
        logging.info(f"Metadata removed from PDF: {pdf_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing PDF {pdf_path}: {e}")
        return False


def remove_metadata_from_docx(docx_path):
    """
    DOCX: Clear core properties with python-docx.
    (ﾉ´ヮ´)ﾉ*:･ﾟ✧
    """
    try:
        logging.info(f"Processing DOCX: {docx_path}")
        doc = Document(docx_path)
        metadata_fields = [
            'author', 'comments', 'category', 'content_status',
            'identifier', 'keywords', 'language', 'last_modified_by',
            'last_printed', 'revision', 'subject', 'title', 'version'
        ]
        for prop in metadata_fields:
            try:
                setattr(doc.core_properties, prop, "")
            except ValueError:
                pass
        doc.settings.odd_and_even_pages_header_footer = False
        doc.save(docx_path)
        logging.info(f"Metadata removed from DOCX: {docx_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing DOCX {docx_path}: {e}")
        return False


def remove_metadata_from_pptx(pptx_path):
    """
    PPTX: Clear core properties with python-pptx.
    ヾ(*ΦωΦ)ツ
    """
    if not CAN_HANDLE_PPTX:
        logging.warning("python-pptx not installed; cannot process PPTX.")
        return False
    try:
        logging.info(f"Processing PPTX: {pptx_path}")
        ppt = Presentation(pptx_path)
        props = ppt.core_properties
        # ~(^-^)~ Clear 'em all
        props.author = ""
        props.category = ""
        props.comments = ""
        props.content_status = ""
        props.created = None
        props.identifier = ""
        props.keywords = ""
        props.last_modified_by = ""
        props.last_printed = None
        props.modified = None
        props.revision = ""
        props.subject = ""
        props.title = ""
        ppt.save(pptx_path)
        logging.info(f"Metadata removed from PPTX: {pptx_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing PPTX {pptx_path}: {e}")
        return False


def remove_metadata_from_ppt(ppt_path):
    """
    PPT (legacy):
    Typically convert .ppt -> .pptx, then do the PPTX route.
    (╯°□°)╯︵ ┻━┻
    """
    try:
        logging.warning(f".ppt is legacy binary; recommended to convert {ppt_path} to .pptx first.")
        return False
    except Exception as e:
        logging.exception(f"Error processing PPT {ppt_path}: {e}")
        return False


def remove_metadata_from_odt(odt_path):
    """
    ODT (OpenDocument Text): Use odfpy to remove <office:meta>.
    (ˆ-ˆ)و♪
    """
    if not CAN_HANDLE_ODF:
        logging.warning("odfpy not installed; cannot process ODT.")
        return False
    try:
        logging.info(f"Processing ODT: {odt_path}")
        doc = odf_load(odt_path)
        meta = doc.meta
        # ~(^-^)~ Remove all metadata children
        for child in list(meta.childNodes):
            meta.removeChild(child)
        doc.save(odt_path)
        logging.info(f"Metadata removed from ODT: {odt_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing ODT {odt_path}: {e}")
        return False


def remove_metadata_from_ods(ods_path):
    """
    ODS (OpenDocument Spreadsheet): Also via odfpy.
    (☞ﾟヮﾟ)☞
    """
    if not CAN_HANDLE_ODF:
        logging.warning("odfpy not installed; cannot process ODS.")
        return False
    try:
        logging.info(f"Processing ODS: {ods_path}")
        spreadsheet = odf_load(ods_path)
        meta = spreadsheet.meta
        for child in list(meta.childNodes):
            meta.removeChild(child)
        spreadsheet.save(ods_path)
        logging.info(f"Metadata removed from ODS: {ods_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing ODS {ods_path}: {e}")
        return False


def remove_metadata_from_epub(epub_path):
    """
    EPUB: Search for .opf files, remove <metadata> content.
    ヾ(〃^∇^)ﾉ
    """
    temp_dir = "temp_epub_extract"
    try:
        logging.info(f"Processing EPUB: {epub_path}")
        os.makedirs(temp_dir, exist_ok=True)
        with zipfile.ZipFile(epub_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # ~(^-^)~ Look for .opf
        opf_files = []
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith('.opf'):
                    opf_files.append(os.path.join(root, file))

        # ~(^-^)~ Strip known metadata tags
        import re
        for opf_file in opf_files:
            try:
                with open(opf_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                # Remove <metadata> sections in a naive way
                content = re.sub(r'<metadata[^>]*>.*?</metadata>', '<metadata></metadata>', content, flags=re.DOTALL)
                with open(opf_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                logging.info(f"Metadata stripped from OPF file: {opf_file}")
            except Exception as ex:
                logging.error(f"Could not strip metadata from {opf_file}: {ex}")

        # ~(^-^)~ Re-zip
        with zipfile.ZipFile(epub_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)
        logging.info(f"Re-zipped EPUB after metadata removal: {epub_path}")

        # ~(^-^)~ Cleanup
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(temp_dir)

        return True
    except Exception as e:
        logging.exception(f"Error processing EPUB {epub_path}: {e}")
        return False


def remove_metadata_from_rtf(rtf_path):
    r"""
    RTF: Naive \info removal with regex.
    (ﾉ´ヮ´)ﾉ*:･ﾟ✧
    """
    try:
        logging.info(f"Processing RTF: {rtf_path}")
        with open(rtf_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        import re
        # Remove {\info ...} blocks
        content = re.sub(r'{\\info[^}]*}', '', content, flags=re.DOTALL)
        with open(rtf_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logging.info(f"Metadata removed from RTF: {rtf_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing RTF {rtf_path}: {e}")
        return False


def remove_metadata_from_mp3(mp3_path):
    """ 
    MP3: Delete ID3 tags with mutagen.
    (＾▽＾) 
    """
    try:
        logging.info(f"Processing MP3: {mp3_path}")
        audio = MP3(mp3_path, ID3=ID3)
        audio.delete()
        audio.save(mp3_path)
        logging.info(f"Metadata removed from MP3: {mp3_path}")
        return True
    except ID3NoHeaderError:
        try:
            audio = MP3(mp3_path)
            audio.save(mp3_path)
            logging.info(f"No ID3 header found, but file saved: {mp3_path}")
            return True
        except Exception as e:
            logging.exception(f"Error saving MP3 {mp3_path} without ID3 header: {e}")
            return False
    except Exception as e:
        logging.exception(f"Error processing MP3 {mp3_path}: {e}")
        return False


def remove_metadata_from_flac(flac_path):
    """ 
    FLAC: Remove tags with mutagen.
    (｡•̀ᴗ-)✧
    """
    try:
        logging.info(f"Processing FLAC: {flac_path}")
        audio = FLAC(flac_path)
        audio.delete()
        audio.save()
        logging.info(f"Metadata removed from FLAC: {flac_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing FLAC {flac_path}: {e}")
        return False


def remove_metadata_from_xlsx(xlsx_path):
    """
    XLSX: Clear workbook properties with openpyxl.
    (∿°○°)∿
    """
    try:
        logging.info(f"Processing XLSX: {xlsx_path}")
        workbook = load_workbook(filename=xlsx_path)
        metadata_fields = [
            'creator', 'title', 'subject', 'description',
            'keywords', 'category', 'comments', 'last_modified_by',
            'company', 'manager'
        ]
        for prop in metadata_fields:
            try:
                setattr(workbook.properties, prop, "")
            except ValueError:
                pass
        workbook.save(xlsx_path)
        logging.info(f"Metadata removed from XLSX: {xlsx_path}")
        return True
    except Exception as e:
        logging.exception(f"Error processing XLSX {xlsx_path}: {e}")
        return False


################################################################
# ~(^-^)~ MASTER SWITCH: Determine file type + call appropriate remover
################################################################

def remove_metadata(file_path):
    """ 
    Decide how to remove metadata based on file extension. 
    (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
    """
    file_extension = os.path.splitext(file_path)[1].lower()

    # ~(^-^)~ Images
    if file_extension in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']:
        return remove_metadata_from_image(file_path)

    # ~(^-^)~ PDFs
    elif file_extension == '.pdf':
        return remove_metadata_from_pdf(file_path)

    # ~(^-^)~ DOCX
    elif file_extension == '.docx':
        return remove_metadata_from_docx(file_path)

    # ~(^-^)~ MP3 / FLAC
    elif file_extension == '.mp3':
        return remove_metadata_from_mp3(file_path)
    elif file_extension == '.flac':
        return remove_metadata_from_flac(file_path)

    # ~(^-^)~ XLSX
    elif file_extension == '.xlsx':
        return remove_metadata_from_xlsx(file_path)

    # ~(^-^)~ ZIP
    elif file_extension == '.zip':
        return remove_metadata_from_zip(file_path)

    # ~(^-^)~ PPTX / PPT
    elif file_extension == '.pptx':
        return remove_metadata_from_pptx(file_path)
    elif file_extension == '.ppt':
        return remove_metadata_from_ppt(file_path)

    # ~(^-^)~ ODT / ODS
    elif file_extension == '.odt':
        return remove_metadata_from_odt(file_path)
    elif file_extension == '.ods':
        return remove_metadata_from_ods(file_path)

    # ~(^-^)~ EPUB
    elif file_extension == '.epub':
        return remove_metadata_from_epub(file_path)

    # ~(^-^)~ RTF
    elif file_extension == '.rtf':
        return remove_metadata_from_rtf(file_path)

    else:
        logging.warning(f"Unsupported file type {file_extension} for {file_path}.")
        return False


################################################################
# ~(^-^)~ GUI / APPLICATION LOGIC
################################################################

class MetadataRemovalApp:
    """
    (づ｡◕‿‿◕｡)づ 
    A Tkinter-based interface for removing metadata from multiple files.
    Parallel processing included for better performance.
    """

    def __init__(self, master):
        """
        Initialize the main Tkinter window and UI elements.
        Σ(＾∀＾) 
        """
        self.master = master
        self.master.title("Metadata Removal Tool - Production Ready & Cute!")
        self.file_paths = []

        # ~(^-^)~ Create a menu bar
        self.menu_bar = tk.Menu(self.master)
        self.master.config(menu=self.menu_bar)

        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Open", command=self.select_files)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit)
        self.menu_bar.add_cascade(label="File", menu=file_menu)

        # ~(^-^)~ Main frame
        self.main_frame = tk.Frame(self.master, padx=10, pady=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # ~(^-^)~ Listbox + Scrollbar
        self.file_listbox = tk.Listbox(self.main_frame, width=80, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.main_frame, command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ~(^-^)~ Button frame
        self.button_frame = tk.Frame(self.master)
        self.button_frame.pack(pady=10)

        self.btn_select_files = tk.Button(self.button_frame, text="Select Files", command=self.select_files)
        self.btn_select_files.grid(row=0, column=0, padx=5)

        self.btn_remove_metadata = tk.Button(self.button_frame, text="Remove Metadata", command=self.process_files)
        self.btn_remove_metadata.grid(row=0, column=1, padx=5)

        self.btn_clear_list = tk.Button(self.button_frame, text="Clear List", command=self.clear_file_list)
        self.btn_clear_list.grid(row=0, column=2, padx=5)

        self.btn_exit = tk.Button(self.button_frame, text="Exit", command=self.master.quit)
        self.btn_exit.grid(row=0, column=3, padx=5)

        # ~(^-^)~ Progress bar
        self.progress = ttk.Progressbar(self.master, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(pady=5)

        # ~(^-^)~ Status label
        self.status_label = tk.Label(self.master, text="No files selected.")
        self.status_label.pack(fill=tk.X, pady=5)

    def select_files(self):
        """ 
        Let user pick multiple files and display them in our listbox.
        ٩(^ᴗ^)۶ 
        """
        new_files = filedialog.askopenfilenames()
        if new_files:
            self.file_paths.extend(new_files)
            self.update_file_listbox()
            self.status_label.config(text=f"{len(self.file_paths)} file(s) selected.")

    def clear_file_list(self):
        """ 
        Clear the list of files and the listbox. 
        (･ω･)つ⊂(･ω･)
        """
        self.file_paths.clear()
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="File list cleared.")

    def update_file_listbox(self):
        """ 
        Refresh the listbox to show all selected file paths.
        (≧▽≦)
        """
        self.file_listbox.delete(0, tk.END)
        for path in self.file_paths:
            self.file_listbox.insert(tk.END, path)

    def process_files(self):
        """
        Remove metadata from each file in parallel. 
        Update progress bar as we go.
        (ﾉ>ω<)ﾉ :｡･::･ﾟ’
        """
        if not self.file_paths:
            messagebox.showerror("Error", "No files selected.")
            return

        self.status_label.config(text="Processing...")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.file_paths)

        def worker(fpath):
            return (fpath, remove_metadata(fpath))

        successes = 0
        with ThreadPoolExecutor() as executor:
            # ~(^-^)~ Submit tasks
            future_map = {executor.submit(worker, fp): fp for fp in self.file_paths}
            completed = 0

            for future in as_completed(future_map):
                fpath, result = future.result()
                completed += 1
                self.progress["value"] = completed
                self.progress.update_idletasks()

                if result:
                    successes += 1
                else:
                    short_name = os.path.basename(fpath)
                    messagebox.showerror("Error", f"Failed to remove metadata from {short_name}")
                    logging.error(f"Failed to remove metadata from {short_name}")

        # ~(^-^)~ Show summary
        if successes == len(self.file_paths):
            messagebox.showinfo("Success", "Metadata removed from all selected files. (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧")
            logging.info("Metadata removed from all selected files.")
        else:
            messagebox.showinfo("Partial Success",
                                f"Metadata removed from {successes} out of {len(self.file_paths)} files.")
            logging.info(f"Partial success: {successes}/{len(self.file_paths)}")

        self.status_label.config(text="Process completed.")


################################################################
# ~(^-^)~ MAIN ENTRY POINT
################################################################

def main():
    """ 
    Initiate the Tkinter loop. 
    ＼(￣▽￣)／
    """
    root = tk.Tk()
    app = MetadataRemovalApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
