from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re

class Elittech(unittest.TestCase):
    def setUp(self):
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", "C:\\Prices\\_downloads")
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk","application/xls,application/octet-stream,application/vnd.ms-excel,application/x-excel,application/x-msexcel,application/excel")
        self.driver = webdriver.Firefox(ffprofile)
        self.driver.implicitly_wait(30)
        self.base_url = "http://elittech.ru/"
        self.verificationErrors = []
        self.accept_next_alert = True
    
    def test_elittech(self):
        driver = self.driver
        driver.get(self.base_url + "/")
        driver.find_element_by_css_selector("div.image").click()
        driver.find_element_by_css_selector("a.download_ico.xls").click()
        time.sleep(20)

    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
