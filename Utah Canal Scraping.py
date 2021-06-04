"""
Created on Sun May 23 17:43:44 2021

@author: vivian
"""

# Imports modules from Selenium that are necessary to send inputs and read data.
# Commented-out modules are not needed, but are provided in case more complex
# implementation is required later on.

from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
# from selenium.webdriver.support import expected_conditions as ec

# Imports the xlsxwriter module to write the collected data to an Excel sheet.
import xlsxwriter as xlsx

# Directory path to the Chrome webdriver file.
# Download the Chrome webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Setup an Excel .xlsx file and formatting information.
            WORKBOOK = xlsx.Workbook("Utah Water Info Worksheet.xlsx")
            companyHeader = WORKBOOK.add_format({'bold' : True, 'bg_color' : '#E6917C'})
            rightHeader = WORKBOOK.add_format({'bold' : True, 'bg_color' : '#5D88F5'})
            linkHeader = WORKBOOK.add_format({'bold' : True, 'bg_color' : '#7CE68A'})
            
            companyFormat = WORKBOOK.add_format({'bg_color' : '#E6917C'})
            rightFormat = WORKBOOK.add_format({'bg_color' : '#5D88F5'})
            tableFormat = WORKBOOK.add_format({'bg_color' : '#82BAFF'})
            linkFormat = WORKBOOK.add_format({'bg_color' : '#7CE68A'})

            # Open main page and pull company URLs from the main table whose 
            # entry lists the county as Uintah.
            driver.get("https://waterrights.utah.gov/canalinfo/canal_owners.asp")
            urlBoxes = driver.find_elements_by_xpath("//tr[td[3]/font[contains(., 'Uintah')]]/td[1]/font/a")
            
            # Sets up worksheets in the Excel file for every company.
            WORKSHEET = [None for _ in range(len(urlBoxes))]
            for index, box in enumerate(urlBoxes):
                WORKSHEET[index] = WORKBOOK.add_worksheet(box.text)
                
            for index, box in enumerate(urlBoxes):
                # Sets up formatting for each company worksheet.
                WORKSHEET[index].set_column(0, 0, 15)
                WORKSHEET[index].set_column(1, 1, 35)
                WORKSHEET[index].set_column(2, 5, 25)
                WORKSHEET[index].set_column(6, 6, 60)
                WORKSHEET[index].set_column(7, 7, 60)
                WORKSHEET[index].set_column(8, 8, 100)
                
                WORKSHEET[index].write(0, 0, "Id Number", companyHeader)
                WORKSHEET[index].write(0, 1, "Company", companyHeader)
                WORKSHEET[index].write(0, 2, "County", companyHeader)
                WORKSHEET[index].write(0, 3, "Water Source", companyHeader)
                WORKSHEET[index].write(0, 4, "Water Right Area", companyHeader)
                
                companyId = box.text
                WORKSHEET[index].write(1, 0, companyId)
                link = box.get_attribute("href")
                
                # Opens company link.
                driver.execute_script("window.open(arguments[0]);", link)
                driver.switch_to.window(driver.window_handles[1])
                
                # Records company information to it's corresponding worksheet.
                try:
                    companyName = driver.find_element_by_xpath("//input[@id='oldCompanyNameId']").get_attribute("value")
                    WORKSHEET[index].write(1, 1,companyName, companyFormat)
                except:
                    pass
                try:
                    companyCounty = driver.find_element_by_xpath("//input[@name='countyName']").get_attribute("value")
                    WORKSHEET[index].write(1, 2, companyCounty, companyFormat)
                except:
                    pass    
                try:
                    companySource = driver.find_element_by_xpath("//input[@id='sourceSaveId']").get_attribute("value")
                    WORKSHEET[index].write(1, 3, companySource, companyFormat)
                except:
                    pass
                try:
                    companyArea = driver.find_element_by_xpath("//table[1]/tbody/tr[11]/td[2]").text
                    WORKSHEET[index].write(1, 4, companyArea, companyFormat)
                except:
                    pass
                
                # Pulls list of water rights from page, if they exist.
                try:
                    waterRights = driver.find_elements_by_xpath("//table[3]/tbody/tr[td[2]/a]")
                        
                    WORKSHEET[index].write(2, 1, "Right ID", rightHeader)
                    WORKSHEET[index].write(2, 2, "Right Status", rightHeader)
                    WORKSHEET[index].write(2, 3, "Priority Date", rightHeader)
                    WORKSHEET[index].write(2, 4, "Quantity (acft)", rightHeader)
                    WORKSHEET[index].write(2, 5, "Flow (cfs)", rightHeader)
                    WORKSHEET[index].write(2, 6, "Source", rightHeader)
                except:
                    pass
                
                # Dummy offset variable to get vertical aligning correct.
                offset = 0
                for rightIndex, right in enumerate(waterRights):
                    
                    # Writes water right info to worksheet.
                    rightNumber = right.find_element_by_xpath("./td[2]/a").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 1, rightNumber, rightFormat)
  
                    rightStatus = right.find_element_by_xpath("./td[4]/span").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 2, rightStatus, rightFormat)

                    rightDate = right.find_element_by_xpath("./td[5]").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 3, rightDate, rightFormat)

                    rightQuantity = right.find_element_by_xpath("./td[6]").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 4, rightQuantity, rightFormat)

                    rightFlow = right.find_element_by_xpath("./td[7]").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 5, rightFlow, rightFormat)
                                    
                    rightSource = right.find_element_by_xpath("./td[8]").text
                    WORKSHEET[index].write(3 + rightIndex + offset, 6, rightSource, rightFormat)
                        
                    # Opens water right link.
                    right.find_element_by_xpath("./td[2]/a").click()
                    driver.switch_to.window(driver.window_handles[2])
                    
                    try:
                        # Pulls 'Points of Diversion' table from page, if it exists.
                        table = wait(driver, 150).until(
                            lambda d:
                                d.find_elements_by_xpath("//tbody[tr/td/span[contains(., 'Points of Diversion')]]/tr")
                            ) 
                    except:
                        table = None
                        pass
                    
                    if table != None:
                        # Writes crude copy of table underneath it's corresponding
                        # water right entry in worksheet.
                        for tableRow in table:
                            tableColumn = tableRow.find_elements_by_xpath("./td")
                            for columnIndex, entry in enumerate(tableColumn):
                                try:
                                    textEntry = entry.find_element_by_xpath("./*").text
                                    if not str.isspace(textEntry):
                                        WORKSHEET[index].write(3 + rightIndex + offset, 2 + columnIndex, textEntry, tableFormat)
                                    else:
                                        WORKSHEET[index].write(3 + rightIndex + offset, 2 + columnIndex, '', tableFormat)
                                except:
                                    try:
                                        textEntry = entry.text
                                        if not str.isspace(textEntry):
                                            WORKSHEET[index].write(3 + rightIndex + offset, 2 + columnIndex, textEntry, tableFormat)
                                        else:
                                            WORKSHEET[index].write(3 + rightIndex + offset, 2 + columnIndex, '', tableFormat)
                                    except:
                                        WORKSHEET[index].write(3 + rightIndex + offset, 2 + columnIndex, '', tableFormat)
                                        pass
                            offset += 1
                    else:
                        WORKSHEET[index].write(4 + rightIndex + offset, 2, "No Table", tableFormat)
                        offset += 1
                    
                    # Selects and open 'Related Documents' option.
                    dropdown = driver.find_element_by_xpath("//select[@id = 'related']/option[2]").get_attribute("value")
                    driver.execute_script("window.open(arguments[0]);", dropdown)
                    driver.switch_to.window(driver.window_handles[3])
                        
                    try:
                        # Sorts documents by date (Newest to Oldest) and record the
                        # download link to the worksheet.
                        driver.find_element_by_xpath("//button[@type = 'submit']").click()
                        driver.find_element_by_xpath("//button[@accesskey = 'P']").click()
                                    
                        driver.switch_to.window(driver.window_handles[4])
                        
                        WORKSHEET[index].write(4 + rightIndex + offset, 2, "Document Link", linkHeader)
                        WORKSHEET[index].write(4 + rightIndex + offset, 3, driver.current_url, linkFormat)
                        offset += 1
                        
                        driver.close()
                        driver.switch_to.window(driver.window_handles[3])
                        
                        driver.close()
                        driver.switch_to.window(driver.window_handles[2])
                    except:
                        # If there are no documents, close the window immediately
                        # and record "No Documents".
                        WORKSHEET[index].write(3 + rightIndex + offset, 8, "No Documents")
                        
                        # Closes the document page and switches to water right page.
                        driver.close()
                        driver.switch_to.window(driver.window_handles[2])
                    
                    # Closes the water right page and switches to company page.
                    driver.close()                                    
                    driver.switch_to.window(driver.window_handles[1])
                
                # Closes company page and returns to main page.
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            
        finally:
            # Save the final .xlsx file to the same directory and closes the browser.
            WORKBOOK.close()
            driver.quit()
        
main()
