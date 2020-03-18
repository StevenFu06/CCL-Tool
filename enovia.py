"""API for Enovia

Date: 2020-03-04
Revision: C
Author: Steven Fu
"""


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import wait, expected_conditions as ec
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

import pickle
import os
import time
import pathlib
from bs4 import BeautifulSoup


class Enovia:
    """Main enovia api, complete with multithreading/ multiprocessing"""

    ENOVIA_URL = 'http://amsnv-enowebp2.netadds.net/enovia/emxLogin.jsp'
    TIMEOUT = 10

    XHOME = '//td[@title="Home"]'
    XSEARCH = '//input[@name="AEFGlobalFullTextSearch"]'
    XSPEC = '//*[contains(text(),"Specifications")]'
    XCATEGORY = '//*[@title="Categories"]'

    def __init__(self, username: str, password: str, headless=True):
        self.username = username
        self.password = password
        self.headless = headless
        self.chrome_options = Options()
        self.browser = None
        self.searched = None

    def __enter__(self):
        self.create_env()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.browser.close()
        self.browser.quit()

    def close(self):
        self.browser.close()
        self.browser.quit()

    def create_env(self):
        """Sets browser and logs in"""

        # Sets browser & options
        if self.headless:
            self.chrome_options.add_argument('--headless')
        self.browser = webdriver.Chrome(options=self.chrome_options)
        # Login
        self.browser.get(self.ENOVIA_URL)
        self.browser.find_element_by_name('login_name').send_keys(self.username)
        temp_pass = self.browser.find_element_by_name('login_password')
        temp_pass.send_keys(self.password)
        temp_pass.send_keys(Keys.ENTER)

    def reset(self):
        """Reset back to enovia startup screen"""
        # For some reason enovia duplicates all elements on startup
        self.browser.get(self.ENOVIA_URL)  # Browser.get to reset and to avoid any preloaded frames

    def search(self, value: str):
        """Searches for the value provided in the search"""

        self.searched = value
        self.reset()
        search_bar = self._wait(ec.element_to_be_clickable((By.XPATH, self.XSEARCH)))
        search_bar.clear()
        search_bar.send_keys(value)
        search_bar.send_keys(Keys.ENTER)
        # Puts the new search window into focus
        self._wait(ec.frame_to_be_available_and_switch_to_it('windowShadeFrame'))

    def open_last_result(self):
        """Opens the last result in the list regardless of any other conditions

        Note:
            Need a better way to open search items. I.e. distinguish between revisions,
            prelim vs proto vs release vs obsolete, etc...
        """
        self._wait(ec.frame_to_be_available_and_switch_to_it('structure_browser'))

        part = self._wait(ec.presence_of_all_elements_located((
            By.XPATH,
            f'//td[@title={self.searched}]//*[contains(text(),"{self.searched}")]'
        )))
        # Javascript click for multiprocessing, normal clicks dont always register
        self.browser.execute_script("arguments[0].click();", part[-1])

        # Brings the selected part number into focus
        self.browser.switch_to.default_content()
        self._wait(ec.frame_to_be_available_and_switch_to_it('content'))
        self._wait(ec.frame_to_be_available_and_switch_to_it('detailsDisplay'))

    def open_latest_state(self, state):
        """Opens the last state in search, i.e. last Released

        Note: Open last result is faster ~2 seconds
        """
        self._wait(ec.frame_to_be_available_and_switch_to_it('structure_browser'))
        # Tests to see if the iframe is loaded before getting soup
        self._wait(ec.presence_of_element_located((By.XPATH, f'//td[@title={self.searched}]')))
        soup = BeautifulSoup(self.browser.page_source, 'html.parser')
        # Catch any typo error/ if the state doesnt exist
        try:
            object_id = soup.find_all('td', {'title': state})[-1]['rmbid']
        except IndexError:
            raise FileNotFoundError(f'{state} for {self.searched} not found')

        part = self.browser.find_element_by_xpath(f'//td[@rmbid="{object_id}"]//a[@class="object"]')
        # Javascript click for multiprocessing, normal clicks dont always register
        self.browser.execute_script("arguments[0].click();", part)

        # Brings the selected part number into focus
        self.browser.switch_to.default_content()
        self._wait(ec.frame_to_be_available_and_switch_to_it('content'))
        self._wait(ec.frame_to_be_available_and_switch_to_it('detailsDisplay'))

    def download_specification_files(self, path):
        """Downloads all files under specifications

        Note:
            There must be a more direct way to download all the files directly instead of
            actually going into specifications.
        """
        self._enable_download_headless(path)
        self._wait(ec.element_to_be_clickable((By.XPATH, self.XCATEGORY))).click()
        self._wait(ec.element_to_be_clickable((By.XPATH, self.XSPEC))).click()
        # Cannot directly click because after every click, enovia refreshes
        # After every click download link gets refreshed so needs to get new every time
        num_downloads = len(self._wait(
            ec.presence_of_all_elements_located((By.XPATH, '//*[@title="Download"]'))
        ))
        for i in range(num_downloads):
            self._wait(
                ec.element_to_be_clickable((By.XPATH, f'(//*[@title="Download"])[{i+1}]'))
            ).click()  # Re-get link based on which is being downloaded now
            ##############
            # NEEDS TO CHANGE, is here because download needs to checkout causing delay
            time.sleep(3)
            ##############
            self.wait_until_downloaded(path)

    def _wait(self, expected_condition):
        """Expected conditions wrapper"""

        try:
            return wait.WebDriverWait(self.browser, self.TIMEOUT).until(expected_condition)
        except TimeoutException:
            raise TimeoutException(f'Timeout at step {expected_condition.__dict__}')

    def _enable_download_headless(self, download_dir):
        """Enables headless download
        I have no idea what this does but it allows headless downloads, and I am forever greatfull
        for the person who has had to sit there and figure this out
        """
        self.browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {
            'cmd': 'Page.setDownloadBehavior',
            'params': {
                'behavior': 'allow',
                'downloadPath': download_dir
            }
        }
        self.browser.execute("send_command", params)

    def wait_until_downloaded(self, download_path):
        """Waits until download is finished before continuing"""

        def is_finished(path):
            all_files = os.listdir(path)
            for file in all_files:
                if pathlib.Path(file).suffix == '.crdownload':
                    return False
            return True

        stahp = time.time() + self.TIMEOUT
        while not is_finished(download_path):
            if time.time() >= stahp:
                raise TimeoutException('Something went wrong while downloading')


if __name__ == '__main__':

    with open('credentials', 'rb') as read:
        cred = pickle.load(read)

    from concurrent.futures import ThreadPoolExecutor

    def multithreading(part):
        enovia = Enovia(cred['user'], cred['pass'], headless=True)
        with enovia as enovia:
            enovia.search(part)
            enovia.open_last_result()
            enovia.download_specification_files(os.getcwd())

    test_list = ['5068482', '5068249', '5068248', '5074524']
    with ThreadPoolExecutor(max_workers=len(test_list)) as executor:
        for result in executor.map(multithreading, test_list):
            pass
