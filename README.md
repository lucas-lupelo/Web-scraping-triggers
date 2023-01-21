# Web-scraping-triggers
This script is a web scraping tool that is designed to extract economic indicators from the website www.tradingeconomics.com. The script is written in Python and uses the Selenium and Beautiful Soup libraries to scrape data from the website, specifically for the United States. The script also uses tkinter to create a message box, openpyxl to edit the XLSX file, os library and pyautogui to open the file.

The script starts by importing necessary libraries including Selenium, Beautiful Soup, tkinter, openpyxl, os and pyautogui. Then it defines a function formatar_num to format the data scraped from the website. Next, the script gets the current date and prompts the user to input a date range for the data to be scraped, it validates the input dates to ensure the format is correct.

After that, the script opens a Chrome browser in headless mode (i.e, it will not be visible to the user) and navigates to the webpage https://tradingeconomics.com/calendar/inflation, it then selects the United States as the country of interest and saves the selection.

After that, the script navigates to the page with the data of interest, scrapes the data and stores it in an excel sheet. This project can be useful for automating data collection and entry for financial analysis or research. It's worth to mention that the script uses different libraries such as Selenium, BeautifulSoup, tkinter, Openpyxl, os and pyautogui to automate the scraping, data handling and data validation.

The script extracts the data and stores it in an XLSX file in a hidden sheet named "Datos_brutos". When the XLSX file is opened, it runs a VBA script that takes the web scraping data and filters and organizes it in the "Data_extract" hidden sheet. From this sheet, the "2022" sheet is updated using VLOOKUP excel function. The script only extracts information from 2022. In February, when data for 2023 starts to be released, the script will extract and consolidate the data from 2022 and then I will create a new sheet for 2023 and update the script accordingly. The "input_from" date will be set to the first date that 2023 data is available, and only the "input_until" date will be prompted to the user, as data from other years is not necessary for this project.

 For bugs or suggestions, please open an issue.
