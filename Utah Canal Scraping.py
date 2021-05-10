# -*- coding: utf-8 -*-
"""
Created on Thu May  6 11:50:23 2021

@author: dania
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import xlsxwriter as xlsx
import time

PATH = "C:\Program Files (x86)\chromedriver.exe"

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            driver.get("https://waterrights.utah.gov/canalinfo/canal_owners.asp")
            WORKBOOK = xlsx.Workbook("Utah Water Info.xlsx")
            WORKSHEET = WORKBOOK.add_worksheet()
            WORKSHEET.write(0, 0, "Id Number")
            WORKSHEET.write(0, 1, "Company")
            WORKSHEET.write(0, 2, "County")
            WORKSHEET.write(0, 3, "Water Source")
            
            urls = driver.find_elements_by_xpath("//tr[td[3]/font[contains(., 'Uintah')]]/td[1]/font/a")

            for index, box in enumerate(urls):
                try:
                    link = box.get_attribute("href")
                    
                    driver.execute_script("window.open(arguments[0]);", link)
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(0.5)
                            
                    # Collect information
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
                            
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                        
                except:
                    pass

        finally:
            WORKBOOK.close()
            driver.quit()
            
main()