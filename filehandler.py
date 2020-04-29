import io
import re
import os
from zipfile import ZipFile
from shutil import copyfile, rmtree, copytree

from enovia import Enovia
from package import Parser, _re_doc_num

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
    poppler = 'Poppler\\bin'

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
            return 'Sch'
        elif assembly_match(text):
            return 'Assy'
        else:
            return None
    except Exception as e:
        text = pdf_to_text_tess(path, tesseract_path)
        if schematic_match(text):
            return 'Sch'
        elif assembly_match(text):
            return 'Assy'
        else:
            return None


class Illustration:

    def __init__(self, ccl, save_dir, processes=1):
        self.ccl = ccl
        self.processes = processes
        self.save_dir = save_dir
        self.filtered = None
        self.scan_dir = None

    def get_filtered(self):
        self.filtered = Parser(self.ccl).filter()

    def _multi_identify(self, pn):
        if self.scan_dir is None:
            raise FileNotFoundError('Scan directory is not given')

        for root, dirs, files in os.walk(os.path.join(self.scan_dir, pn)):
            for file in files:
                if file.endswith('.pdf'):
                    src = os.path.join(root, file)
                    result = identify(src)
                    print(f'Identified {pn} - {file} to be {result}')
                    if result is not None:
                        dnum = _re_doc_num(file)[0]
                        dest = os.path.join(self.save_dir, f'{pn}-{result}. {dnum}.pdf')
                        copyfile(src, dest)

    def get_illustrations(self, scan_dir=None):
        """Automatically scan folder (scan_dir) for illustrations

        Note:
            In order for scan dir to work properly folders must be in the form
            Part number
                - Files
                - Files
                - Files
            The same format that the DocumentCollector temp creator is in
        """
        if self.scan_dir is None and scan_dir is None:
            raise FileNotFoundError('Scan directory is not given')
        elif scan_dir is not None:
            self.scan_dir = scan_dir
        # Pandas dataframe uses numpy 64 floats wich automatically add a .0
        pns = [pn.replace('.0', '') for pn in self.filtered['pn'].astype(str)]
        if self.filtered is None:
            self.get_filtered()
        pool = Pool(self.processes)
        pool.map(self._multi_identify, pns)

        self._used, self._count = [], 0
        for idx in self.filtered.index:
            self._rename(idx)

    def _rename(self, idx):
        """Function to renumber the illustrations in the folder"""

        for file in os.listdir(self.save_dir):
            pn = self.filtered.loc[idx, 'pn'].astype(str).replace('.0', '')
            try:
                ill_type = file.split("-")[1]
                if file.split('-')[0] == pn and pn not in self._used:
                    self._used.append(ill_type)
                    self._count += 1
                    src = os.path.join(self.save_dir, file)
                    renamed = f'Ill.{self._count} {pn} {self.filtered.loc[idx, "desc"]} {ill_type}'
                    # Convert to windows accepable name
                    renamed = re.sub(r'[^a-zA-Z0-9()#.]+', ' ', renamed)
                    dest = os.path.join(self.save_dir, renamed)
                    os.rename(src, dest)
            except IndexError:
                continue

    def update_ccl(self):
        """Updates the word ccl"""
        pass

    def insert_illustration(self):
        """Inserts a new illustration into the ccl, requires UI"""
        pass

    def delete_illustration(self):
        """Delets an illustration from the ccl, requires UI"""
        pass


class DocumentCollector:

    def __init__(self, username, password, ccl, save_dir, processes=1, headless=True):
        self.username = username
        self.password = password
        self.ccl = ccl
        self.filtered = None
        self.save_dir = save_dir
        self.processes = processes
        self.failed = []
        self.headless = headless
        self.temp_dir = None

    def create_temp_dir(self):
        self.temp_dir = os.path.join(self.save_dir, 'temp')
        if os.path.exists(self.temp_dir):
            self.clear_temp()
        os.makedirs(self.temp_dir)

    def get_filtered(self):
        self.filtered = Parser(self.ccl).filter()

    def _multidownload(self, pn: str):
        temp_path = os.path.join(self.temp_dir, pn)
        try:
            print(f'{pn} is downloading')
            if not os.path.exists(temp_path):
                os.makedirs(temp_path)
            with Enovia(self.username, self.password, headless=self.headless) as enovia:
                enovia.search(pn)
                enovia.open_last_result()
                enovia.download_specification_files(temp_path)
            print(f'{pn} has downloaded sucessfully')
            return None
        except:
            print(f'{pn} failed to download')
            return pn

    def download(self):
        try:
            rmtree(self.temp_dir)
        except TypeError:
            pass
        self.create_temp_dir()

        if self.filtered is None:
            self.get_filtered()
        pns = self.filtered['pn'].astype(str)
        pns = [pn.replace('.0', '') for pn in pns]

        pool = Pool(self.processes)
        self.failed = [failed for failed in pool.map(self._multidownload, pns) if failed is not None]

    def extract_all(self):
        """Extracts alll the files into the main part folder removing any zip files"""

        # Regex check for vendor zip files
        def vendor(string):
            result = re.findall(r'VENDOR', string, re.IGNORECASE)
            return True if result else False

        # Extract all pdf
        for root, dirs, files in os.walk(self.temp_dir):
            for file in files:
                if file.endswith('.zip'):
                    with ZipFile(os.path.join(root, file)) as zip_file:
                        zip_file.extractall(root)
                    # Clean up
                    os.remove(os.path.join(root, file))
        # Rescan all documents to remove all sub folders and place into main
        for part in os.listdir(self.temp_dir):
            # Only go through directories
            for sub_root, sub_dirs, sub_files in os.walk(os.path.join(self.temp_dir, part)):
                # Scan through sub dirs only
                for sub_dir in sub_dirs:
                    for file in os.listdir(os.path.join(sub_root, sub_dir)):
                        src = os.path.join(sub_root, sub_dir, file)
                        dest = os.path.join(sub_root, file)
                        # Special condition of vendor files, only extract the pdf
                        # Vendor and archive files will remain in the folder
                        if vendor(file):
                            with ZipFile(src) as zip_file:
                                for file_in_zip in zip_file.namelist():
                                    if file_in_zip.endswith('.pdf'):
                                        zip_file.extract(file_in_zip, sub_root)
                    # Clean up
                        copyfile(src, dest)
                    rmtree(os.path.join(sub_root, sub_dir))

    @staticmethod
    def _format_name(pn, desc, fn):
        pn = pn.astype(str).replace('.0', '')
        fn = fn.astype(str).replace('.0', '')
        return pn, desc, fn

    def collect_documents(self, path=None):
        if path is None:
            path = self.save_dir
        path_bold = path
        for idx in self.filtered.index:
            pn, desc, fn = self._format_name(self.filtered.loc[idx, "pn"],
                                             self.filtered.loc[idx, "desc"],
                                             self.filtered.loc[idx, "fn"])
            folder_name = f'{pn} {desc} (#{fn})'
            folder_name = re.sub(r"[^a-zA-Z0-9()#]+", ' ', folder_name)
            temp_folder = os.path.join(self.temp_dir, pn)

            try:
                if self.filtered.loc[idx, 'bold']:
                    path_bold = os.path.join(path, folder_name)
                    copytree(temp_folder, path_bold)

                elif not self.filtered.loc[idx, 'bold']:
                    sub_path = os.path.join(path_bold, folder_name)
                    copytree(temp_folder, sub_path)
            except FileExistsError:
                continue

    def clear_temp(self):
        if self.temp_dir is None:
            raise ValueError('temp_dir is not set')
        if not os.path.exists(self.temp_dir):
            raise FileNotFoundError('Temp directory doesnt exist')
        rmtree(self.temp_dir)


if __name__ == '__main__':
    test = Parser('ccl.docx')
    print(test.filter())
    # document = DocumentCollector('Steven.Fu', 'hipeople1S', 'ccl.docx', 'ccl documents', processes=4)
    # document.download()
    # print(document.failed)
    # document.extract_all()
    # document.collect_documents('collected')
