# Scrape-listings-from-Airbnb
This python program will search for a location on Airbnb website (in english) and copy all the listing details to an excel spreadsheet using Selenium.
It's really useful for who wants to have a better visualization of all the places available.

Main steps of this code:

1) Set preferences that will disable a bunch of functionalities of the browser (loading images, for example), making your code faster
2) Access the location on Airbnb website using the Firefox Webdriver
3) Copy all the links of the accommodations
4) Access each link and copy the information to an excel spreadsheet using openpyxl and BeautifulSoup
