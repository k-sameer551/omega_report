import time, os
from selenium import webdriver
from selenium.webdriver import EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from omega.constants import Constants as const
from omega.utils import Utils

class Omega(webdriver.Edge):
    """omega class file inherting webdriver.edge"""

    def __init__(self, driver_path= const.DRIVER_PATH, teardown=False):
        self.driver_path = driver_path
        self.teardown = teardown
        documents_path = Utils.get_download_path()
        options = EdgeOptions()
        prefs = {'download.default_directory' : documents_path}
        options.add_experimental_option('detach', True)
        options.add_experimental_option('prefs', prefs)
        options.add_argument('start_maximized')
        options.add_argument('disable-popup-blocking')
        options.add_argument('disable-infobars')
        super(Omega, self).__init__(options=options)
        self.implicitly_wait(30)
        

    def __exit__(self, exc_type, exc, traceback):
        """exit browser window when true"""
        if self.teardown:
            self.quit()
    

    def land_web_page(self):
        """land on desired web page"""
        self.get(const.URL_PATH)
        time.sleep(5)
        self.find_element(By.XPATH, '//*[@id="login"]').click()
        time.sleep(2)
        self.find_element(By.XPATH, '/html/body/div[6]/button').click()
        time.sleep(5)
        self.get(const.NAVIGATE_PATH)
        Select(self.find_element(By.ID, 'BusinessArea')).select_by_visible_text('CSS UNET')
        Select(self.find_element(By.ID, 'TimeZone')).select_by_visible_text('Central Standard Time')


    def download_report(self, report_type, new_file_name, download_path, timeout):
        """download type of report"""
        select = Select(self.find_element(By.ID, 'SelectedWorkTypeId'))
        select.select_by_visible_text(report_type)
        time.sleep(5)
        self.find_element(By.ID, 'cloneBtn').click()
        time.sleep(5)
        self.find_element(By.ID, 'btnSubmit').click()
        time.sleep(30)
        self.switch_to.frame('riframe')
        links = self.find_elements(By.TAG_NAME, "a")
        for link in links:
            if link.get_attribute("title") == 'Excel':
                self.execute_script("arguments[0].click();",link)
                break
        time.sleep(30)
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(download_path)
            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds +=1
        self.switch_to.default_content()
        Utils.rename_file(new_file_name, download_path)
        return True