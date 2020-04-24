import io
import re
import os
from shutil import copyfile

from enovia import Enovia
from package import Parser

from multiprocessing import Pool

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
    pages = pdf2image.convert_from_path(path, dpi=resolution, grayscale=True, poppler_path=poppler)

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


class DocumentCollector:

    def __init__(self, username, password, ccl, save_dir, processes=1):
        self.username = username
        self.password = password
        self.ccl = ccl
        self.filtered = None
        self.save_dir = save_dir
        self.processes = processes

    def get_filtered(self):
        self.filtered = Parser(self.ccl).filter()

    def multidownload(self, pn: str):
        temp_path = os.path.join(self.save_dir, 'temp', pn)
        os.makedirs(temp_path)
        with Enovia(self.username, self.password, headless=True) as enovia:
            enovia.search(pn)
            try:
                enovia.open_latest_state('Prototype')
            except FileNotFoundError:
                enovia.open_latest_state('Released')
            enovia.download_specification_files(temp_path)

    def download(self):
        os.makedirs(os.path.join(self.save_dir, 'temp'))
        if self.filtered is None:
            self.get_filtered()
        pns = self.filtered['pn'].astype(str)
        pns = [pn.replace('.0', '') for pn in pns]

        pool = Pool(self.processes)
        pool.map(self.multidownload, pns)


def collect_documents(username, password, ccl_word, save_dir):
    def format_name(pn, desc, fn):
        pn = pn.astype(str).replace('.0', '')
        fn = fn.astype(str).replace('.0', '')
        return pn, desc, fn

    enovia = Enovia(username, password, headless=True)
    filtered = Parser(ccl_word).filter()
    path_bold = save_dir

    for idx in filtered.index:
        pn, desc, fn = format_name(filtered.loc[idx, "pn"], filtered.loc[idx, "desc"], filtered.loc[idx, "fn"])
        folder_name = f'{pn} {desc} (#{fn})'

        if filtered.loc[idx, 'bold'] == True:
            path_bold = os.path.join(save_dir, folder_name)

        elif filtered.loc[idx, 'bold'] == False:
            sub_path = os.path.join(path_bold, folder_name)


if __name__ == '__main__':
    documents = DocumentCollector('test', 'test', 'ccl.docx', os.getcwd())
    documents.download()
