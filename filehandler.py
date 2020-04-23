import io
import re
import os

from multiprocessing import Pool

from shutil import copyfile

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import pytesseract as pt

import pdf2image


def pdf_to_text_miner(path):
    pdf_manager = PDFResourceManager()
    with io.StringIO() as temp_file:

        with TextConverter(pdf_manager, temp_file, laparams=LAParams()) as converter:
            interpreter = PDFPageInterpreter(pdf_manager, converter)

            with open(path, 'rb') as read:
                for page in PDFPage.get_pages(read):
                    interpreter.process_page(page)

        return temp_file.getvalue()


def pdf_to_text_tess(path, tesseract_path, resolution=250):
    # Set tesseract path
    pt.pytesseract.tesseract_cmd = tesseract_path
    poppler = 'Popper\\bin'

    # Read pdf as image
    pages = pdf2image.convert_from_path(path, dpi=resolution, grayscale=True, poppler=poppler)

    # Extract text using Google's tesseract
    text = [pt.image_to_string(page, lang='eng') for page in pages]

    del pages
    return ' '.join(text)


def schematic_match(string):
    sciex = re.findall(r'Sciex', string, re.IGNORECASE)
    schem = re.findall(r'SCHEM\*', string)
    return True if sciex and schem else False


def assembly_match(string):
    sciex = re.findall(r'Sciex', string, re.IGNORECASE)
    schem = re.findall(r'SCHEM\*', string)
    # Projection seems to be to hard find since the word is too long
    projection = re.findall(r'projection', string, re.IGNORECASE)
    scale = re.findall(r'scale', string, re.IGNORECASE)
    desc = re.findall(r'PART DESCRIPTION', string, re.IGNORECASE)
    return True if sciex and (scale or projection or desc) and not schem else False


def identify(path):
    """Identify the type of file the pdf is (sch, assy or neither)"""

    tesseract_path = 'Tesseract-OCR\\tesseract.exe'
    try:
        text = pdf_to_text_miner(path)
        # Added an additional check to see if no text was picked up by the miner
        if len(text) < 500:
            raise Exception
        elif schematic_match(text):
            return 'schematic'
        elif assembly_match(text):
            return 'assembly'
        else:
            return 'none'
    except Exception as e:
        text = pdf_to_text_tess(path, tesseract_path)
        print(f'exception "{e}" happened had to use tesseract')
        if schematic_match(text):
            return 'schematic'
        elif assembly_match(text):
            return 'assembly'
        else:
            return 'none'


def multiprocess(document):
    current_path = document[0]
    file = document[1]
    try:
        if file[-3:] == 'pdf':
            print(f'Scanning {file}')
            file_path = os.path.join(current_path, file)
            result = identify(file_path)
            print(f'Identified to be {result}')
            print()
            dest_path = os.path.join(result, file)
            copyfile(file_path, os.path.join('identify', dest_path))
    except:
        copyfile(os.path.join(current_path, file), os.path.join('identify', os.path.join('failed', file)))


if __name__ == '__main__':
    import time

    path = 'Illustrations'
    ccl_docs = os.walk('CCL Documents')

    documents = []
    for root, dir, files in ccl_docs:
        for file in files:
            documents.append((root, file))

    pool = Pool(12)

    start = time.time()
    pool.map(multiprocess, documents)
    end = time.time()

    print(f'scanning {len(documents)} documents took {end - start}s')
