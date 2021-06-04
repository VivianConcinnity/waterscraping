"""
Created on Sun May 30 19:44:17 2021

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

# Directory path to Chrome webdriver file.
# Download the Chrome webdriver at https://chromedriver.chromium.org/downloads
PATH = "C:\Program Files (x86)\chromedriver.exe"

# NRCS sites to navigate to.
# Add or remove links as needed.
URLS = ["https://wcc.sc.egov.usda.gov/nwcc/site?sitenum=833", "https://wcc.sc.egov.usda.gov/nwcc/site?sitenum=559"]

def main():
    with webdriver.Chrome(PATH) as driver:
        try:
            # Opens dummy page to instantiate a new active session. . .
            driver.get("https://www.google.com")
            for link in URLS: 
                # . . .and opens new window with desired link from URLS.
                driver.execute_script("window.open(arguments[0]);", link)
                driver.switch_to.window(driver.window_handles[1])
                
                # Selects list of options for report content and for historic year.
                # "wait(driver, 120)" is used to give time for every element to load.
                reportContent = wait(driver, 120).until(lambda d: d.find_elements_by_xpath("//select[@name = 'report']/option")) 
                historicYears = driver.find_elements_by_xpath("//select[@name = 'year']/option")
                
                # Selects daily option for Time Series.
                driver.find_element_by_xpath(
                    "//select[@name = 'timeseries']/option[1]"
                    ).click()
                
                # Selects .csv option for Format.
                driver.find_element_by_xpath(
                    "//select[@name = 'format']/option[@value = 'copy']"
                    ).click()
                
                # Selects Historic Calendar year.
                driver.find_element_by_xpath(
                    "//select[@name = 'month']/option[@value = 'CY']"
                    ).click()
                
                # Selects Daily for day option.
                driver.find_element_by_xpath(
                    "//select[@name = 'day']/option[1]"
                    ).click()
                
                # Downloads required Historic calendary year report.
                for option in reportContent:
                    # Checks if current option is an unwanted option.
                    if not ((option.get_attribute("value") in ["ALL"]) or (option.text == "===Individual elements===")):                
                        option.click()
                        for year in historicYears:
                            year.click()
                            driver.find_element_by_xpath(
                                "//input[@class = 'scanReportButtonGreen']"
                                ).click()
                
                # Closes current site.
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            
        finally:
            # Waits a set time to allow downloads to finish, then closes the browser.
            # Adjust time if needed according to download speed.
            time.sleep(30)
            driver.quit()
            

main()
