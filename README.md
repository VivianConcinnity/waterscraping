# Water Info Webscraping Project

This repository holds three files that were used to collect information on water rights and water canals in the Utah area from various websites. The code is provided for transparency, and this document is provided to give some insight as to how the code works and what parameters can be modified.

At the top of every Python file (after the importing of packages) are various variables whose names are in all-caps. These variables are parameters, and they can be changed safely to change the behaviour of the program.

## Utah Canal Scraping

The `Utah Canal Scraping.py` file is used to pull certain water canal information from the [Utah Division of Water Rights](https://waterrights.utah.gov/canalinfo/canal_owners.asp) site and writes it to an Excel .xlsx file. This data includes:

- ID Number
- Company Name
- County Name
- Water Source
- Water Right Information
  - Water Right ID
  - Water Right Status
  - Priority Date
  - Quantity (acre-feet)
  - Flow (cfs)
  - Water Right Source
  - Points of Diversion
- Link to a compiled PDF of documents, if documents exist

### Parameters

The `PATH` variable is the directory path to the Chrome webdriver file on the computer the Python file is run on. 
Change this variable to the appropriate directory where the `chromedriver.exe` file is downloaded to.

The `COUNTY_NAME` variable tells the program what to check in the County column of the table, and pulls the entry's link if it matches the County name.
Change this variable to another valid county's name from the same page.
Currently, the program is pulling info from the Uintah county's canals.

The `FILE_NAME` variable is the name of the .xlsx file saved to the same directory as the Python file.
Change the variable to the desired name you want to save the .xlsx file under.

## Station Webscraping

The `Station Webscraping.py` file navigates to the [Ashley Creek system page](https://www.waterrights.utah.gov/distribution/WaterRecords.asp?system_name=ASHLEY%20CREEK) on the Utah Division of Water Rights site, pulls information from each station in the Ashley Creek system, and writes it to a spreadsheet in an Excel .xlsx file. This data includes:

- Station ID Number
- Link to a compiled PDF of documents, if documents exist
- Raw ASCII table of Daily/Monthly/Yearly Summary Table
- Comma-delimited table of Daily/Monthly/Yearly Summary Table
  - For both of the above tables, the table for the longest unit of time is pulled and recorded, if available (Yearly first, followed by Monthly and Daily).

### Parameters

The `PATH` variable is the directory path to the Chrome webdriver file on the computer the Python file is run on. Change this variable to the appropriate directory where the `chromedriver.exe` file is downloaded to.

The `SYSTEM_URL` variable is the system page's URL to which the program navigates to.
Change the URL to another valid system page URL from the same website to pull info for a different county.
Currently, the program is pulling info from the Ashley Creek system.

The `FILE_NAME` variable is the name of the .xlsx file saved to the same directory as the Python file.
Change the variable to the desired name you want to save the .xlsx file under.

## NRCS Webscraping

The `NRCS Webscraping.py` file navigates to the [Kings Cabin](https://wcc.sc.egov.usda.gov/nwcc/site?sitenum=559) and [Trout Creek](https://wcc.sc.egov.usda.gov/nwcc/site?sitenum=833) website on the Natural Resources Conservation Service website and automatically downloads all daily calendar-year .csv report files from the website.

These files include:

- Standard SNOTEL
- Soil Moisture & Temperature
- Soil Moisture
- Soil & Air Temperature
- Accumulated Precipitation
- Accumulated Precipitation & Snow
- Air Temperature
- Precipitation Accumulation
- Snow Depth
- Snow Water Equivalent
- Soil Moisture Percent
- Soil Temperature

### Parameters

The `PATH` variable is the directory path to the Chrome webdriver file on the computer the Python file is run on. 
Change this variable to the appropriate directory where the `chromedriver.exe` file is downloaded to.

The `URLS` variable is an array of SNOTEL water site URLs from the Natural Resources Conversation Service website to which the program navigates to.
Add or remove SNOTEL site URLs to the array that you want to download reports from.

The `DOWNLOAD_WAIT` variable is used to make the program wait at the end of execution to allow all the files to download.
Change the number as needed to give the program enough time for all the files to download; the exact number will depend on the strength of the Wi-Fi signal the computer the code is running on. A poor connection will require a larger number to wait for all files to download; a faster connection can use a smaller number to make the program terminate faster (though, in either case, it's best to err on using a larger number to allow time for all files to download).

The `FILE_NAME` variable is the name of the .xlsx file saved to the same directory as the Python file.
Change the variable to the desired name you want to save the .xlsx file under.
