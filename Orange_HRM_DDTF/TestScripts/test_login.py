"""
Perform the Login of Orange HRM
1. Using Excel File
"""
# Import all Necessary Dependencies
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

#Import Explicitly Wait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


#Use Page Object Model Import Data and Locators
from TestData.data import Data
from TestLocators.locators import Locators
from Utilities.excel_functions import Excel_Function

class TestOrangeDDTF:
    def test_login_excel(self):
        # binding the file path and sheet number
        self.excel_file = Data.excel_file
        self.sheet_number = Data.sheet_number

        # create excel object passing the file and sheet number
        self.excel = Excel_Function(self.excel_file, self.sheet_number)

        # Initialize Driver
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        # launch the browser and maximize window
        self.driver.get(Data.url)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 10)
        self.driver.implicitly_wait(10)

        # get the total row values
        self.rows = self.excel.row_count()

        # using for loop to iterate row by row values
        for row in range(2, self.rows + 1):
            username = self.excel.read_data(row, 5)
            password = self.excel.read_data(row, 6)

            # enter the username read from the excel file
            self.driver.find_element(by=By.NAME, value=Locators.username_locator).send_keys(username)

            # enter the password read from the excel file
            self.driver.find_element(by=By.NAME, value=Locators.password_locator).send_keys(password)

            # click the login button
            login_click = self.wait.until(EC.element_to_be_clickable((By.XPATH,Locators.login_button_locator)))
            login_click.click()

            # Main validation logic of login testing (PASS/FAIL)
            if Data.dashboard_url in self.driver.current_url:
                print("SUCCESS: Login successful")
                self.excel.write_data(row, 7, "TEST PASS")
                self.driver.find_element(by=By.XPATH,value=Locators.drop_down_selector).click()
                self.driver.find_element(by=By.XPATH,value=Locators.logout).click()


            elif Data.url in self.driver.current_url:
                print("FAIL: Login failed")
                self.excel.write_data(row, 7, "TEST FAIL")
                self.driver.refresh()

        self.driver.quit()






