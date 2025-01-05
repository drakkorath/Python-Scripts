import os
import time
import zipfile
import logging
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from fpdf import FPDF

# Configure logging
logging.basicConfig(
    filename="file_monitor.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
)


class PDFConverter:
    def __init__(self, pdf_file_path):
        self.pdf_file_path = pdf_file_path

    def images_to_pdf(self, images):
        pdf = FPDF()
        for image in images:
            pdf.add_page()
            pdf.image(image, x=10, y=10, w=190)
        pdf.output(self.pdf_file_path)
        logging.info(f"Created PDF: {self.pdf_file_path}")


class MyHandler(FileSystemEventHandler):
    def __init__(self, zip_file_path):
        super().__init__()
        self.zip_file_path = zip_file_path

    def on_any_event(self, event):
        if event.is_directory:
            return
        elif event.src_path == self.zip_file_path and (
            event.event_type == "created" or event.event_type == "modified"
        ):
            logging.info(f"File {event.src_path} has been {event.event_type}")
            if os.path.exists(event.src_path):
                try:
                    self.process_event()
                except Exception as e:
                    logging.error(f"Error processing event: {e}")

    def process_event(self):
        base_dir = os.path.abspath(
            os.path.dirname(__file__)
        )  # Get the directory of the script
        temp_dir = os.path.join(base_dir, "atemp")
        pdf_path = os.path.join(base_dir, "photos.pdf")

        try:
            extract_zip(self.zip_file_path, temp_dir)
            delete_file(self.zip_file_path)
            create_pdf_from_directory(temp_dir, pdf_path)
            clean_up_temp_directory(temp_dir)
        except Exception as e:
            logging.error(f"Error processing event: {e}")


def extract_zip(zip_file, extract_to):
    """Extracts the contents of a ZIP file to a specified directory."""
    with zipfile.ZipFile(zip_file, "r") as zObject:
        zObject.extractall(path=extract_to)


def delete_file(file_path):
    """Deletes a file if it exists."""
    if os.path.exists(file_path):
        os.remove(file_path)
        logging.info(f"Deleted file: {file_path}")
    else:
        logging.warning(f"File {file_path} does not exist.")


def create_pdf_from_directory(directory, output_pdf):
    """Creates a PDF from images in a specified directory."""
    pdf = FPDF()
    for filename in os.listdir(directory):
        if filename.endswith(".jpg"):
            filepath = os.path.join(directory, filename)
            if os.path.isfile(filepath):
                pdf.add_page()
                pdf.image(filepath, x=10, y=10, w=190)
    pdf.output(output_pdf)
    logging.info(f"Created PDF: {output_pdf}")


def clean_up_temp_directory(directory):
    """Deletes all files and subdirectories within a directory."""
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
                logging.info(f"Deleted file: {file_path}")
            elif os.path.isdir(file_path):
                clean_up_temp_directory(file_path)
                os.rmdir(file_path)
                logging.info(f"Deleted directory: {file_path}")
        except Exception as e:
            logging.error(f"Failed to delete {file_path}: {e}")


class PDFHandler(FileSystemEventHandler):
    def __init__(self, excel_path, download_dir):
        self.excel_path = excel_path
        self.download_dir = download_dir

    def on_created(self, event):
        if event.is_directory:
            return
        file_path, file_extension = os.path.splitext(event.src_path)
        if file_extension.lower() == ".pdf":
            self.rename_pdf(event.src_path)

    def rename_pdf(self, pdf_path):
        # Read the excel file
        df = pd.read_excel(self.excel_path)

        if df.empty:
            logging.warning("excel file is empty. No rows to process.")
            return
        # Get the first row's data
        first_row = df.iloc[0]
        new_name = first_row.iloc[0]  # Use iloc for positional indexing

        # Sanitize the new name by removing invalid characters for filenames
        new_name = "".join(
            c for c in new_name if c.isalnum() or c in (" ", ".", "_")
        ).rstrip()

        # Rename the PDF file
        new_pdf_path = os.path.join(self.download_dir, f"{new_name}.pdf")
        os.rename(pdf_path, new_pdf_path)
        logging.info(f"Renamed PDF to: {new_pdf_path}")

        # Delete the first row and save the excel
        df = df.drop(index=0)
        df.to_excel(self.excel_path, index=False)


if __name__ == "__main__":
    # Monitoring for ZIP files
    monitor_path = r"/your/directory/for/monitoring/goes/here"
    zip_file_path = os.path.join(monitor_path, "download.zip") # downloaded xlsx / csv filename goes here
    pdf_file_path = os.path.join(monitor_path, "photos.pdf") # converted to pdf filename goes here
    event_handler_zip = MyHandler(zip_file_path)
    observer_zip = Observer()
    observer_zip.schedule(event_handler_zip, path=monitor_path, recursive=False)

    # Monitoring for PDF files to rename
    excel_path = r"/your/directory/for/excel/file/goes/here"
    event_handler_pdf = PDFHandler(excel_path, monitor_path)
    observer_pdf = Observer()
    observer_pdf.schedule(event_handler_pdf, path=monitor_path, recursive=False)

    # Start both observers
    observer_zip.start()
    observer_pdf.start()
    logging.info(f"Started monitoring {monitor_path}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer_zip.stop()
        observer_pdf.stop()
    observer_zip.join()
    observer_pdf.join()
