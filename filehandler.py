import io
import re
import os
from zipfile import ZipFile
from shutil import copyfile, rmtree, copytree

from enovia import Enovia
from package import Parser, _re_doc_num, _re_pn
import progressbar

from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
from concurrent.futures import as_completed

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import pytesseract as pt
import pdf2image

from selenium.common.exceptions import SessionNotCreatedException


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
        # self.document = Document(ccl)
        # self.table = self.document.tables[0]
        self.processes = processes
        self.save_dir = save_dir
        self.filtered = None
        self.scan_dir = None
        self.ccl_dir = None

    def get_filtered(self):
        self.filtered = Parser(self.ccl).filter()

    def _multi_identify_scan(self, pn):
        if self.scan_dir is None:
            raise FileNotFoundError('Scan directory is not given')
        output_message = []
        for file in os.listdir(os.path.join(self.scan_dir, pn)):
            if file.endswith('.pdf'):
                src = os.path.join(self.scan_dir, pn, file)
                result = identify(src)
                output_message.append(f'Identified {pn} - {file} to be {result}')
                print(f'Identified {pn} - {file} to be {result}')
                if result is not None:
                    dnum = _re_doc_num(file)[0]
                    dest = os.path.join(self.save_dir, f'{pn}-{result}. {dnum}.pdf')
                    copyfile(src, dest)
        return output_message

    def _multi_identify_ccl(self, pn: str):
        if self.ccl_dir is None:
            raise FileNotFoundError('Scan directory is not given')
        output_message = []
        for root, dirs, files in os.walk(self.ccl_dir):
            for dir in dirs:
                found_pn = _re_pn(dir)
                if found_pn == int(pn):
                    for file in os.listdir(os.path.join(root, dir)):
                        if file.endswith('.pdf'):
                            src = os.path.join(root, dir, file)
                            result = identify(src)
                            output_message.append(f'Identified {pn} - {file} to be {result}')
                            print(f'Identified {pn} - {file} to be {result}')
                            if result is not None:
                                dnum = _re_doc_num(file)[0]
                                dest = os.path.join(self.save_dir, f'{pn}-{result}. {dnum}.pdf')
                                copyfile(src, dest)
        return output_message

    def get_illustrations(self, scan_dir=None, ccl_dir=None):
        """Automatically scan folder (scan_dir) for illustrations

        Note:
            In order for scan dir to work properly folders must be in the form
            Part number
                - Files
                - Files
                - Files
            The same format that the DocumentCollector temp creator is in
        """
        multiscan = self._multi_identify_ccl
        if self.scan_dir is None and scan_dir is None and self.ccl_dir is None and ccl_dir is None:
            raise FileNotFoundError('Scan directory is not given')
        elif scan_dir is not None:
            self.scan_dir = scan_dir
            multiscan = self._multi_identify_scan
        elif ccl_dir is not None:
            self.ccl_dir = ccl_dir
            multiscan = self._multi_identify_ccl

        if self.filtered is None:
            self.get_filtered()
        # Pandas dataframe uses numpy 64 floats which automatically add a .0
        pns = [pn.replace('.0', '') for pn in self.filtered['pn'].astype(str)]

        increment = 1/len(pns)
        with ProcessPoolExecutor(max_workers=self.processes) as executor:
            message = [
                executor.submit(multiscan, pn)
                for pn in pns
            ]
            for future in as_completed(message):
                messages = future.result()
                for out in messages:
                    # Demo only commented
                    # progressbar.add_current(increment)
                    print(out)
        progressbar.add_current(1)
        # pool = Pool(self.processes)
        # pool.map(multiscan, pns)

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

    def shift_up_ill(self, shift_from):
        """Shifts illustrations up starting from and including shift_from

        Will also update the CCL
        """
        for file in os.listdir(self.save_dir):
            if file.endswith('.pdf'):
                ill_num = int(re.findall(r'\d+', file.split(' ')[0])[0])
                if ill_num >= shift_from:
                    new_name = file.replace(file.split(' ')[0], f'Ill.{ill_num+1}')
                    dest = os.path.join(self.save_dir, new_name)
                    src = os.path.join(self.save_dir, file)
                    os.rename(src, dest)

    def shift_down_ill(self, shift_from):
        """Shifts illustrations numbers down starting from and including shift_from

        Will also update the CCL
        """
        for file in os.listdir(self.save_dir):
            if file.endswith('.pdf'):
                ill_num = int(re.findall(r'\d+', file.split(' ')[0])[0])
                if ill_num >= shift_from:
                    new_name = file.replace(file.split(' ')[0], f'Ill.{ill_num - 1}')
                    dest = os.path.join(self.save_dir, new_name)
                    src = os.path.join(self.save_dir, file)
                    os.rename(src, dest)


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
        self.progress_val = 0

    def create_temp_dir(self):
        self.temp_dir = os.path.join(self.save_dir, 'temp')
        if os.path.exists(self.temp_dir):
            self.clear_temp()
        os.makedirs(self.temp_dir)

    def get_filtered(self):
        self.filtered = Parser(self.ccl).filter()

    def _multidownload(self, pn: str):
        temp_path = os.path.join(self.temp_dir, pn)
        progressbar.add_current(self.progress_val)
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
        except SessionNotCreatedException as e:
            raise e
        except Exception as e:
            print(f'{pn} failed to download due to {e}')
            return pn

    def download(self, pns):
        """Download from enovia using multi download multiprocess"""

        self.progress_val = 1/len(pns)
        with ThreadPoolExecutor(self.processes) as executor:
            self.failed = [failed for failed in executor.map(self._multidownload, pns) if failed is not None]
        # pool = Pool(self.processes)
        # self.failed = [failed for failed in pool.map(self._multidownload, pns) if failed is not None]
        prev_failed_len = -1
        self.progress_val = 0
        # Rerun until self.failed length becomes constant or is empty
        while self.failed and prev_failed_len != len(self.failed):
            prev_failed_len = len(self.failed)
            with ThreadPoolExecutor(self.processes) as executor:
                self.failed = [failed for failed in executor.map(self._multidownload, self.failed) if failed is not None]
        return self.failed

    def extract_all(self):
        """Extracts alll the files into the main part folder removing any zip files"""

        # Regex check for vendor zip files
        def vendor(string):
            result = re.findall(r'VENDOR', string, re.IGNORECASE)
            return True if result else False

        # For Progress bar increments
        def progress_increment(temp_dir):
            increment = 0
            for root, dirs, files in os.walk(self.temp_dir):
                for file in files:
                    if file.endswith('.zip'):
                        increment += 1
            increment += len(os.listdir(temp_dir))
            return 1/increment
        increment = progress_increment(self.temp_dir)

        # Extract all pdf
        for root, dirs, files in os.walk(self.temp_dir):
            for file in files:
                if file.endswith('.zip'):
                    ## Progress bar ##
                    progressbar.add_current(increment)
                    ## Progress bar ##
                    with ZipFile(os.path.join(root, file)) as zip_file:
                        zip_file.extractall(root)
                    # Clean up
                    os.remove(os.path.join(root, file))
        # Rescan all documents to remove all sub folders and place into main
        for part in os.listdir(self.temp_dir):
            ## Progress bar ##
            progressbar.add_current(increment)
            ## Progress bar ##
            # Only go through directories
            for sub_root, sub_dirs, sub_files in os.walk(os.path.join(self.temp_dir, part)):
                # Scan through sub dirs only
                for sub_dir in sub_dirs:
                    files = (file for file in os.listdir(os.path.join(sub_root, sub_dir))
                             if os.path.isfile(os.path.join(sub_root, sub_dir, file)))
                    for file in files:
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

    def structure(self, path=None):
        """Turns temp structured into proper CCL strucutre

        Param path: the save location of the ccl structured folder
        """
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

    @staticmethod
    def _pn_exists(pn, dirs):
        """Check if pn exists in files"""

        for dir in dirs:
            found = _re_pn(dir)
            if found == int(pn):
                return dir
        return False

    def _check_path_and_copy(self, pn, path, dest_folder):
        """Checks if the path contains a folder with pn and copy contents to dest_folder"""

        for root, dirs, files in os.walk(path):
            dir_found = self._pn_exists(pn, dirs)
            if dir_found is not False:
                # Get Files only
                files = (file for file in os.listdir(os.path.join(root, dir_found))
                         if os.path.isfile(os.path.join(root, dir_found, file)))
                for file in files:
                    src = os.path.join(root, dir_found, file)
                    dest = os.path.join(dest_folder, file)
                    if not os.path.exists(dest_folder):
                        os.makedirs(dest_folder)
                    copyfile(src, dest)
                return True
        return False

    def _check_paths(self, pn, paths, dest_folder):
        for path in paths:
            copied = self._check_path_and_copy(pn, path, dest_folder)
            if copied:
                return True
        return False

    def collect_documents(self, check_paths: list = None):
        """Main function for this class, collect the documents for ccl

        param check_paths: paths to check before downloading off Enovia, index 0 gets highest priority
        """
        try:
            rmtree(self.temp_dir)
        except TypeError:
            pass
        self.create_temp_dir()

        if self.filtered is None:
            self.get_filtered()
        if check_paths is None:
            check_paths = []

        pns = self.filtered['pn'].astype(str)
        pns = [pn.replace('.0', '') for pn in pns]

        to_download = []
        for pn in pns:
            dest_folder = os.path.join(self.temp_dir, pn)
            copied = self._check_paths(pn, check_paths, dest_folder)
            if not copied:
                to_download.append(pn)

        print('Beginning Download')
        self.download(to_download)
        print('Extracting all Zip files')
        self.extract_all()
        print('Structuring Temporary Dire ctory')
        self.structure()
        print('Cleaning up')
        self.clear_temp()
        print(f'{self.failed} have failed to download')

    def clear_temp(self):
        if self.temp_dir is None:
            raise ValueError('temp_dir is not set')
        if not os.path.exists(self.temp_dir):
            raise FileNotFoundError('Temp directory doesnt exist')
        rmtree(self.temp_dir)


if __name__ == '__main__':
    import importlib_metadata
    print(importlib_metadata.version("jsonschema"))
