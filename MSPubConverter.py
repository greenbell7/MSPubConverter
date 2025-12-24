import os
import re
import logging
import win32com.client
from win32com.client import constants, gencache
from tkinter import Tk, filedialog
from tqdm import tqdm

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")

def pick_folder(title):
    root = Tk()
    root.withdraw()  # Hide the main Tk window
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder

def to_camel_case(s: str) -> str:
    # Split on any non-alphanumeric characters, drop empties
    parts = [p for p in re.split(r'[^A-Za-z0-9]+', s) if p]
    if not parts:
        return ''
    first = parts[0].lower()
    rest = ''.join(p.capitalize() for p in parts[1:])
    return first + rest

def unique_filename_no_spaces(output_dir: str, base_name: str, ext: str = ".pdf") -> str:
    """
    Return a filename in output_dir using base_name + ext.
    If that file exists, append -2, -3, ... to base_name until unique.
    """
    candidate = base_name + ext
    counter = 2
    # Use os.path.exists (case-insensitive on Windows) to detect collisions
    while os.path.exists(os.path.join(output_dir, candidate)):
        candidate = f"{base_name}-{counter}{ext}"
        counter += 1
    return candidate

def convert_pub_to_pdf(publisher, input_pub, output_pdf):
    input_pub = os.path.normpath(os.path.abspath(input_pub))
    output_pdf = os.path.normpath(os.path.abspath(output_pdf))
    doc = None
    try:
        logging.info("Opening: %s", input_pub)
        doc = publisher.Open(input_pub)
        try:
            opened_name = getattr(doc, "FullName", None)
            opened_path = getattr(doc, "Path", None)
            logging.info("Opened doc FullName=%s Path=%s", opened_name, opened_path)
        except Exception:
            logging.exception("Unable to read opened document properties")

        if not doc or (isinstance(opened_name, str) and os.path.normcase(opened_name) != os.path.normcase(input_pub)):
            logging.error("Publisher did not open the requested file; opened=%r expected=%r", opened_name, input_pub)
            return False

        # Ensure target directory exists
        out_dir = os.path.dirname(output_pdf)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        logging.info("Attempting ExportAsFixedFormat: %s", output_pdf)
        try:
            # Prefer typelib constants (EnsureDispatch populates them). Fall back to numeric literals.
            fmt_pdf = getattr(constants, "pbFixedFormatTypePDF", 32)
            intent_print = getattr(constants, "pbFixedFormatIntentPrint", 1)
            # Use the common 3-arg form: (Type, Filename, Intent)
            doc.ExportAsFixedFormat(fmt_pdf, output_pdf, intent_print)
        except Exception as ex:
            logging.exception("ExportAsFixedFormat failed: %s", ex)
            try:
                logging.error("Exception args: %r", ex.args)
            except Exception:
                pass
            return False

        # Verify file was created
        if not os.path.exists(output_pdf):
            logging.error("Export completed but output file not found: %s", output_pdf)
            return False

        logging.info("Export succeeded: %s", output_pdf)
        return True
    except Exception:
        logging.exception("ERROR converting %s", input_pub)
        return False
    finally:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                logging.exception("Error closing document for %s", input_pub)

def convert_all_pub_files(parent_folder, output_root):
    # Collect all .pub files first
    pub_files = []
    for root, dirs, files in os.walk(parent_folder):
        for file in files:
            if file.lower().endswith(".pub"):
                pub_files.append(os.path.join(root, file))

    if not pub_files:
        print("No .pub files found.")
        return

    publisher = None
    try:
        # Use EnsureDispatch to generate/attach the typelib wrappers and constants
        try:
            publisher = gencache.EnsureDispatch("Publisher.Application")
        except Exception:
            # fallback to dynamic Dispatch if EnsureDispatch fails
            publisher = win32com.client.Dispatch("Publisher.Application")

        # Optionally hide UI
        try:
            publisher.Visible = False
        except Exception:
            pass

        for input_path in tqdm(pub_files, desc="Converting files", unit="file"):
            relative_path = os.path.relpath(os.path.dirname(input_path), parent_folder)
            if relative_path == ".":
                output_dir = output_root
            else:
                output_dir = os.path.join(output_root, relative_path)
            output_dir = os.path.normpath(output_dir)
            os.makedirs(output_dir, exist_ok=True)

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            camel = to_camel_case(base_name)
            if not camel:
                camel = "file"

            # Generate a unique, no-spaces camelCase filename
            pdf_name = unique_filename_no_spaces(output_dir, camel, ".pdf")
            output_path = os.path.join(output_dir, pdf_name)
            output_path = os.path.normpath(output_path)

            success = convert_pub_to_pdf(publisher, input_path, output_path)
            if not success:
                logging.warning("Failed to convert: %s", input_path)
    finally:
        if publisher is not None:
            try:
                publisher.Quit()
            except Exception:
                logging.exception("Error quitting Publisher application")

if __name__ == "__main__":
    print("=== Microsoft Publisher → PDF Converter ===")

    input_folder = pick_folder("Select the parent folder containing .pub files")
    if not input_folder:
        print("No input folder selected.")
        exit()

    output_folder = pick_folder("Select the output folder for converted PDFs")
    if not output_folder:
        print("No output folder selected.")
        exit()

    convert_all_pub_files(input_folder, output_folder)

    print("\nConversion complete.")
