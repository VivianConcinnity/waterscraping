"""
Created on Thu May  6 11:50:23 2021
@author: vivian
"""

# Imports modules from Selenium that are necessary to send inputs and read data.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# Imports the xlsxwriter module to write the collected data to an Excel sheet.
import xlsxwriter as xlsx

# Imports the time module to implement waits for pages to load and for debugging.
import time

# PATH stores the directory path to the Chrome webdriver file on the host's harddrive.
# Download the Chrome webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Setup an Excel .xlsx file and format columns and rows.
            WORKBOOK = xlsx.Workbook("Utah Water Info.xlsx")
            WORKSHEET = WORKBOOK.add_worksheet()
            
            bold = WORKBOOK.add_format({'bold' : True})

            WORKSHEET.set_column(1, 1, 35)
            WORKSHEET.set_column(2, 5, 25)
            WORKSHEET.write(0, 0, "Id Number", bold)
            WORKSHEET.write(0, 1, "Company", bold)
            WORKSHEET.write(0, 2, "County", bold)
            WORKSHEET.write(0, 3, "Water Source", bold)
            WORKSHEET.write(0, 4, "Water Right Area", bold)
            
            # Open main page and pull URLs from the main table whose 
            # corresponding entry lists Uintah as the county.
            driver.get("https://waterrights.utah.gov/canalinfo/canal_owners.asp")
            urls = driver.find_elements_by_xpath("//tr[td[3]/font[contains(., 'Uintah')]]/td[1]/font/a")
            index = 1
            
            for box in urls:
                try:
                    # For every URL from the table, write the company id to
                    # the spreadsheet. . .
                    link = box.get_attribute("href")
                    companyId = box.text
                    WORKSHEET.write(index, 0, companyId)
                    
                    # . . .open and switch to a new window that navigates
                    # to the URL. . .
                    driver.execute_script("window.open(arguments[0]);", link)
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(0.5)
                            
                    # . . .and collect the required information from the page.
                    # Try/except blocks are used in case the information isn't present on the page.
                    try:
                        companyName = driver.find_element_by_xpath("//input[@id='oldCompanyNameId']").get_attribute("value")
                        WORKSHEET.write(index, 1, companyName)
                    except:
                        pass
                            
                    try:
                        companyCounty = driver.find_element_by_xpath("//input[@name='countyName']").get_attribute("value")
                        WORKSHEET.write(index, 2, companyCounty)
                    except:
                        pass
                            
                    try:
                        companySource = driver.find_element_by_xpath("//input[@id='sourceSaveId']").get_attribute("value")
                        WORKSHEET.write(index, 3, companySource)
                    except:
                        pass
                    
                    try:
                        companyArea = driver.find_element_by_xpath("//table[1]/tbody/tr[11]/td[2]").text
                        WORKSHEET.write(index, 4, companyArea)
                    except:
                        pass

                    try:
                        # If there are water rights on the page, . .
                        #/html/body/div[4]/form/table[3]/tbody/tr[7]
                        waterRights = driver.find_elements_by_xpath("//table[3]/tbody/tr[6]/following-sibling::tr")
                        
                        WORKSHEET.write(index + 1, 1, "Right ID", bold)
                        WORKSHEET.write(index + 1, 2, "Right Status", bold)
                        WORKSHEET.write(index + 1, 3, "Quantity (cfs)", bold)
                        WORKSHEET.write(index + 1, 4, "Flow (acre-feet)", bold)
                        WORKSHEET.write(index + 1, 5, "Source", bold)
                        
                        # . . . write the related information in an indented block underneath the company entry in the Excel spreadsheet.
                        # TODO - First row is being copied repeatedly
                        for jndex, right in enumerate(waterRights):
                            try:
                                rightNumber = right.find_element_by_xpath("//td[2]/a").text
                                rightStatus = right.find_element_by_xpath("//td[4]/span").text
                                rightQuantity = right.find_element_by_xpath("//td[6]").text
                                rightFlow = right.find_element_by_xpath("//td[7]").text
                                rightSource = right.find_element_by_xpath("//td[8]").text

                                WORKSHEET.write(index + 2 + jndex, 1, rightNumber)
                                WORKSHEET.write(index + 2 + jndex, 2, rightStatus)
                                WORKSHEET.write(index + 2 + jndex, 3, rightQuantity)
                                WORKSHEET.write(index + 2 + jndex, 4, rightFlow)
                                WORKSHEET.write(index + 2 + jndex, 5, rightSource)

                                # CLicks link and opens water right window.
                                right.find_element_by_xpath("//td[2]/a").click()
                                driver.switch_to.window(driver.window_handles[2])
                                
                                # TODO - Write code to collect info from new window.
                                time.sleep(2)
                                
                                driver.close()
                                driver.switch_to.window(driver.window_handles[1])
                            except:
                                continue
                            
                        index += len(waterRights) + 3
                    except:
                        index += 3
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
