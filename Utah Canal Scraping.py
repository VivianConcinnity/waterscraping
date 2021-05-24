"""
Created on Sun May 23 17:43:44 2021

@author: vivian
"""

# Imports modules from Selenium that are necessary to send inputs and read data.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

# Imports the xlsxwriter module to write the collected data to an Excel sheet.
import xlsxwriter as xlsx

# Imports the time module to implement waits for pages to load and for debugging.
import time

# PATH stores the directory path to the Chrome webdriver file on the host's hard drive.
# Download the Chrome webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Setup an Excel .xlsx file and format columns and rows.
            WORKBOOK = xlsx.Workbook("Utah Water Info.xlsx")
            WORKSHEET = WORKBOOK.add_worksheet()
            
            bold = WORKBOOK.add_format({'bold' : True})
            
            WORKSHEET.set_column(0, 0, 15)
            WORKSHEET.set_column(1, 1, 35)
            WORKSHEET.set_column(2, 5, 25)
            WORKSHEET.set_column(6, 6, 60)
            WORKSHEET.set_column(7, 7, 60)
            WORKSHEET.set_column(8, 8, 100)
            
            # Open main page and pull URLs from the main table whose 
            # corresponding entry lists Uintah as the county.
            driver.get("https://waterrights.utah.gov/canalinfo/canal_owners.asp")
            urls = driver.find_elements_by_xpath("//tr[td[3]/font[contains(., 'Uintah')]]/td[1]/font/a")
            index = 0
            
            for box in urls:
                try:
                    offset = index + 1
                    WORKSHEET.write(index, 0, "Id Number", bold)
                    WORKSHEET.write(index, 1, "Company", bold)
                    WORKSHEET.write(index, 2, "County", bold)
                    WORKSHEET.write(index, 3, "Water Source", bold)
                    WORKSHEET.write(index, 4, "Water Right Area", bold)
                    
                    # For every URL from the table, write the company id to
                    # the spreadsheet. . .
                    link = box.get_attribute("href")
                    companyId = box.text
                    WORKSHEET.write(index + 1, 0, companyId)
                    
                    # . . .open and switch to a new window that navigates
                    # to the URL. . .
                    driver.execute_script("window.open(arguments[0]);", link)
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(0.5)
                            
                    # . . .and collect the required information from the page.
                    # Try/except blocks are used in case the information isn't present on the page.
                    try:
                        companyName = driver.find_element_by_xpath("//input[@id='oldCompanyNameId']").get_attribute("value")
                        WORKSHEET.write(offset, 1, companyName)
                    except:
                        pass
                            
                    try:
                        companyCounty = driver.find_element_by_xpath("//input[@name='countyName']").get_attribute("value")
                        WORKSHEET.write(offset, 2, companyCounty)
                    except:
                        pass
                            
                    try:
                        companySource = driver.find_element_by_xpath("//input[@id='sourceSaveId']").get_attribute("value")
                        WORKSHEET.write(offset, 3, companySource)
                    except:
                        pass
                    
                    try:
                        companyArea = driver.find_element_by_xpath("//table[1]/tbody/tr[11]/td[2]").text
                        WORKSHEET.write(offset, 4, companyArea)
                    except:
                        pass

                    try:
                        # If there are water rights on the page, . .
                        waterRights = driver.find_elements_by_xpath("//table[3]/tbody/tr[td[2]/a]")
                        
                        WORKSHEET.write(offset + 1, 1, "Right ID", bold)
                        WORKSHEET.write(offset + 1, 2, "Right Status", bold)
                        WORKSHEET.write(offset + 1, 3, "Priority Date", bold)
                        WORKSHEET.write(offset + 1, 4, "Quantity (acft)", bold)
                        WORKSHEET.write(offset + 1, 5, "Flow (cfs)", bold)
                        WORKSHEET.write(offset + 1, 6, "Source", bold)
                        WORKSHEET.write(offset + 1, 7, "Points of Diversion", bold)
                        WORKSHEET.write(offset + 1, 8, "Document Link", bold)
                        
                        # . . . write the related information in an indented block underneath the company entry in the Excel spreadsheet.
                        for jndex, right in enumerate(waterRights):
                            try:
                                try:
                                    rightNumber = right.find_element_by_xpath("./td[2]/a").text
                                    WORKSHEET.write(offset + 2 + jndex, 1, rightNumber)
                                
                                    rightStatus = right.find_element_by_xpath("./td[4]/span").text
                                    WORKSHEET.write(offset + 2 + jndex, 2, rightStatus)
                                
                                    rightDate = right.find_element_by_xpath("./td[5]").text
                                    WORKSHEET.write(offset + 2 + jndex, 3, rightDate)
                                    
                                    rightQuantity = right.find_element_by_xpath("./td[6]").text
                                    WORKSHEET.write(offset + 2 + jndex, 4, rightQuantity)
                        
                                    rightFlow = right.find_element_by_xpath("./td[7]").text
                                    WORKSHEET.write(offset + 2 + jndex, 5, rightFlow)
                                    
                                    rightSource = right.find_element_by_xpath("./td[8]").text
                                    WORKSHEET.write(offset + 2 + jndex, 6, rightSource)
                                except:
                                    pass

                                # CLicks link and opens water right window.
                                try:
                                    right.find_element_by_xpath("./td[2]/a").click()
                                    driver.switch_to.window(driver.window_handles[2])
                                
                                # TODO - Write code to collect info in the instance there are
                                # multiple diversion points.
                                    
                                    wait = WebDriverWait(driver, 120).until(ec.presence_of_element_located((By.XPATH, "//table")))
                                    try:
                                        diversionPoints = driver.find_elements_by_xpath("//div[6]/table/tbody/*/td[@class = 'tblData']/a")
                                        string = ""
                                        for points in diversionPoints:        
                                            string += points.text + " "
                                        WORKSHEET.write(offset + 2 + jndex, 7, string)
                                    except:
                                        # diversionPoint =
                                        # WORKSHEET.write(offset + 2 + jndex, 7, diversionPoint)
                                        pass

                                    dropdown = driver.find_element_by_xpath("//select[@id = 'related']/option[2]").get_attribute("value")
                                    driver.execute_script("window.open(arguments[0]);", dropdown)
                                    driver.switch_to.window(driver.window_handles[3])
                                    try:
                                        driver.find_element_by_xpath("//button[@type = 'submit']").click()
                                        driver.find_element_by_xpath("//button[@accesskey = 'P']").click()
                                    
                                        driver.switch_to.window(driver.window_handles[4])
                                        WORKSHEET.write(offset + 2 + jndex, 8, driver.current_url)
                                        driver.close()
                                        driver.switch_to.window(driver.window_handles[3])
                                    except:
                                        WORKSHEET.write(offset + 2 + jndex, 8, "No Documents")
                                        driver.close()
                                        driver.switch_to.window(driver.window_handles[3])
                                        pass
                                        
                                    driver.close()
                                    
                                    driver.switch_to.window(driver.window_handles[2])
                                    driver.close()
                                    
                                    driver.switch_to.window(driver.window_handles[1])
                                except:
                                    driver.switch_to.window(driver.window_handles[2])
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[1])
                                    pass
                            except:
                                continue
                            
                        index += len(waterRights) + 4
                    except:
                        index += 4
                        pass
                    
                    # After collecting information, the driver closes the window and switches back to the main window.
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                        
                except:
                    pass

        finally:
            # After everything is done, close and save the .xlsx file to the
            # same directory and close the browser.
            WORKBOOK.close()
            driver.quit()
            
main()
