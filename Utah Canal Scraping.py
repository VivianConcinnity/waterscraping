# -*- coding: utf-8 -*-
"""
Created on Thu May  6 11:50:23 2021

@author: vivian
"""

# Imports modules from selenium that are necessary to send inputs and read data.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# Imports the xlsxwriter module to write data to an Excel sheet.
import xlsxwriter as xlsx

# Imports the time module to implement waits for pages to load.
import time

# PATH stores the directory path to the Chrome webdriver file on the host's drive.
# Download the webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Setup an Excel .xlsx file and format columns and rows.
            WORKBOOK = xlsx.Workbook("Utah Water Info.xlsx")
            WORKSHEET = WORKBOOK.add_worksheet()
            WORKSHEET.write(0, 0, "Id Number")
            WORKSHEET.write(0, 1, "Company")
            WORKSHEET.write(0, 2, "County")
            WORKSHEET.write(0, 3, "Water Source")
            
            # Open main page. . .
            driver.get("https://waterrights.utah.gov/canalinfo/canal_owners.asp")
            
            # . . .and pull URLs from table that list Uintah as the county.
            urls = driver.find_elements_by_xpath("//tr[td[3]/font[contains(., 'Uintah')]]/td[1]/font/a")

            for index, box in enumerate(urls):
                try:
                    # For every URL from the table. . .
                    link = box.get_attribute("href")
                    
                    # . . .opens and switches to a new window that navigates to the URL. . .
                    driver.execute_script("window.open(arguments[0]);", link)
                    driver.switch_to.window(driver.window_handles[1])
                    
                    # . . .wait half a second to let any time-sensitive elements load. . .
                    time.sleep(0.5)
                            
                    # . . .and collect the required information from the page.
                    # Try/except blocks are used in case the information isn't present on the page.
                    try:
                        companyId = driver.find_element_by_xpath("//input[@id='oldCompanyNumberId']").get_attribute("value")
                        WORKSHEET.write(index + 1, 0, companyId)
                    except:
                        pass
                    
                    try:
                        companyName = driver.find_element_by_xpath("//input[@id='oldCompanyNameId']").get_attribute("value")
                        WORKSHEET.write(index + 1, 1, companyName)
                    except:
                        pass
                            
                    try:
                        county = driver.find_element_by_xpath("//input[@name='countyName']").get_attribute("value")
                        WORKSHEET.write(index + 1, 2, county)
                    except:
                        pass
                            
                    try:
                        waterSource = driver.find_element_by_xpath("//input[@id='sourceSaveId']").get_attribute("value")
                        WORKSHEET.write(index + 1, 3, waterSource)
                    except:
                        pass
                    
                    # After collecting information, the driver closes the window and switches back to the main window.
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                        
                except:
                    pass

        finally:
            # After everything is done, close and save the .xlsx file and close the browser.
            WORKBOOK.close()
            driver.quit()
            
main()
