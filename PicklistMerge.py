import os
from PyPDF2 import PdfMerger
import datetime


def combine_pdf_files(pdf_directory_path, output_directory_path):
    """
    Combines all PDF files in the specified directory into a single PDF file.

    Parameters:
    - pdf_directory_path: The path to the directory containing the PDF files to be combined.
    - output_directory_path: The path to the directory where the combined PDF file will be saved.
    """
    current_datetime = datetime.datetime.now().strftime('%d_%m_%Y_%H%M')
    output_filename = f"Picklist{current_datetime}.pdf"
    combined_pdf_path = os.path.join(output_directory_path, output_filename)

    # Create a PdfMerger object
    merger = PdfMerger()

    # Collect all PDF files in the given directory
    pdf_file_paths = [os.path.join(pdf_directory_path, f) for f in os.listdir(pdf_directory_path) if f.endswith('.pdf')]

    # Loop through the collected PDF file paths and append them to the merger object
    for pdf_path in pdf_file_paths:
        merger.append(pdf_path)

    # Write out the combined PDF
    merger.write(combined_pdf_path)
    merger.close()

    print(f"All PDF files in {pdf_directory_path} have been combined into {combined_pdf_path}")
