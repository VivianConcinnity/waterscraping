# -*- coding: utf-8 -*-
"""
Created on Tue Jun  8 10:01:20 2021

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
import time

# Imports the xlsxwriter module to write the collected data to an Excel sheet.
import xlsxwriter as xlsx

# Directory path to Chrome webdriver file.
# Download the Chrome webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"


def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Opens distribution water records page.
            driver.get("https://www.waterrights.utah.gov/distribution/WaterRecords.asp?system_name=ASHLEY%20CREEK")
            
            # Pulls entries from table.
            stationEntries = wait(driver, 120).until(
                    lambda d:
                        d.find_elements_by_xpath("//tr[td/a]")
                )
            
                
            # Sets up Excel spreadsheet with appropriate number of sheets.
            WORKBOOK = xlsx.Workbook("Station Distribution Records.xlsx")
            WORKSHEET = [None for _ in range(len(stationEntries))]
            
            # Sets up cell formatting options.
            stationFormat = WORKBOOK.add_format({'bg_color' : '#E6917C'})
            tableFormat = WORKBOOK.add_format({'bg_color' : '#82BAFF'})
            tableASCIIFormat = WORKBOOK.add_format({'bg_color' : '#82BAFF', 'font_name' : 'Courier New'})
            linkFormat = WORKBOOK.add_format({'bg_color' : '#7CE68A'})

            for entryIndex, entry in enumerate(stationEntries):
                stationName = entry.find_element_by_xpath("./td[2]").text
                
                # Trims station name if the name is too long for the worksheet to store.
                if len(stationName) > 25 - (len(str(entryIndex + 1))):
                    stationName = stationName[:25 - (len(str(entryIndex + 1)))] + "... (" + str(entryIndex + 1) +")"
                else:
                    stationName = stationName[:25] + "... (" + str(entryIndex + 1) + ")"
                    
                # Adds new worksheet with name of station.
                WORKSHEET[entryIndex] = WORKBOOK.add_worksheet(stationName)
                
                # Opens and switches to station info window.
                stationLink = entry.find_element_by_xpath("./td/a").get_attribute("href")
                driver.execute_script("window.open(arguments[0]);", stationLink)
                driver.switch_to.window(driver.window_handles[1])
                
                # Pulls text from page.
                pageText = driver.find_element_by_xpath("//pre").text
                linesOfText = pageText.split("\n")
                
                # Pulls station ID from url and writes to Excel document.
                tempURL = driver.current_url
                stationID = tempURL[str.find(tempURL, "ID=") + 3: str.find(tempURL, "&R")]
                WORKSHEET[entryIndex].write(0, 0, "Station ID:   " + stationID, stationFormat)
                
                # Writes page text to Excel document.
                for lineIndex, textLine in enumerate(linesOfText):
                    textLine = textLine.strip()
                    WORKSHEET[entryIndex].set_column(0, 0, 155)
                    WORKSHEET[entryIndex].write(lineIndex + 1, 0, textLine, stationFormat)
                
                driver.find_element_by_xpath("//input[@type = 'BUTTON']").click()
                
                # Adjust time as needed to let new page load.
                time.sleep(1)
                
                driver.switch_to.window(driver.window_handles[2])
                
                offset = len(linesOfText)
                
                # Gets link to compiled documents, if there are any.
                try:
                    driver.find_element_by_xpath("//button[@type = 'submit']").click()
                    driver.find_element_by_xpath("//button[@accesskey = 'P']").click()
                    
                    driver.switch_to.window(driver.window_handles[3])
                    documentLink = driver.current_url
                    
                    WORKSHEET[entryIndex].write(offset + 2, 0, "Document Link: " + documentLink, linkFormat)
                    
                    driver.close()
                    driver.switch_to.window(driver.window_handles[2])
                except:
                    # Writes "No Documents" if no documents are available.
                    WORKSHEET[entryIndex].write(offset + 2, 0, "Document Link: No Documents", linkFormat)
                
                driver.close()
                driver.switch_to.window(driver.window_handles[1])
                
                # Opens raw ASCII table page.
                driver.find_element_by_xpath("//select/option[@value = 'Daily_Sum' or @value = 'Monthly_Ascii' or @value = 'Annual_Ascii']").click()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[2])
                
                # Writes raw ASCII table to Excel document.
                try:
                    element = wait(driver, 5).until(
                        lambda d: d.find_element_by_xpath("//pre[not(select)]")
                        )
                finally:
                    summaryTable = element.text.split("\n")
                    for tableIndex, row in enumerate(summaryTable):
                        WORKSHEET[entryIndex].write(offset + 4 + tableIndex, 0, row, tableASCIIFormat)
                    
                    # Closes ASCII page.
                    driver.close()
                    driver.switch_to.window(driver.window_handles[1])
                
                # Opens comma-delimited page.
                driver.find_element_by_xpath("//select/option[@value = 'Daily_Comma' or @value = 'Monthly_Comma' or @value = 'Annual_Comma']").click()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[2])
                
                try:
                    element = wait(driver, 5).until(
                        lambda d: 
                            d.find_element_by_xpath("(//a[@href[substring(.,string-length(.) - string-length('txt') + 1) = 'txt']])[last()]")
                        )
                finally:
                    try:
                        # Opens new page of comma-delimited values.
                        link = element.get_attribute("href")
                        driver.execute_script("window.open(arguments[0]);", link)
                        driver.switch_to.window(driver.window_handles[3])
                        
                        # Formats array of table data. . .
                        pageValues = driver.find_element_by_xpath("//pre").text
                        tableOfValues = pageValues.split("\n")
                        for tableIndex, row in enumerate(tableOfValues):
                            temp = row.split(",")
                            tableOfValues[tableIndex] = temp
                        
                        # . . .and writes it to the Excel spreadsheet.
                        for rowIndex, entry in enumerate(tableOfValues):
                            for columnIndex, text in enumerate(entry):
                                WORKSHEET[entryIndex].write(rowIndex, 2 + columnIndex, text.strip("\" "), tableFormat)
                        
                        # Closes comma-delimited page
                        driver.close()
                        driver.switch_to.window(driver.window_handles[2])
                    except:
                        pass
            
                driver.close()
                driver.switch_to.window(driver.window_handles[1])
                
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        finally:
            # After everything is done, close the browser.
            print(driver.current_url)
            WORKBOOK.close()
            driver.quit()
            
main()