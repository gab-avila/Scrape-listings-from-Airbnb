from bs4 import BeautifulSoup
import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import re
from openpyxl.styles import Font
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import sys

#Set english language
profile = webdriver.FirefoxProfile()
profile.set_preference('intl.accept_languages', 'en-US, en')

#MAKE SELENIUM FASTER
profile.set_preference("network.http.pipelining", True)
profile.set_preference("network.http.proxy.pipelining", True)
profile.set_preference("network.http.pipelining.maxrequests", 8)
profile.set_preference("content.notify.interval", 500000)
profile.set_preference("content.notify.ontimer", True)
profile.set_preference("content.switch.threshold", 250000)
profile.set_preference("browser.cache.memory.capacity", 65536) # Increase the cache capacity.
profile.set_preference("browser.startup.homepage", "about:blank")
profile.set_preference("reader.parse-on-load.enabled", False) # Disable reader, we won't need that.
profile.set_preference("browser.pocket.enabled", False) # Duck pocket too!
profile.set_preference("loop.enabled", False)
profile.set_preference("browser.chrome.toolbar_style", 1) # Text on Toolbar instead of icons
profile.set_preference("browser.display.show_image_placeholders", False) # Don't show thumbnails on not loaded images.
profile.set_preference("browser.display.use_document_colors", False) # Don't show document colors.
profile.set_preference("browser.display.use_document_fonts", 0) # Don't load document fonts.
profile.set_preference("browser.display.use_system_colors", True) # Use system colors.
profile.set_preference("browser.formfill.enable", False) # Autofill on forms disabled.
profile.set_preference("browser.helperApps.deleteTempFileOnExit", True) # Delete temprorary files.
profile.set_preference("browser.shell.checkDefaultBrowser", False)
profile.set_preference("browser.startup.homepage", "about:blank")
profile.set_preference("browser.startup.page", 0) # blank
profile.set_preference("browser.tabs.forceHide", True) # Disable tabs, We won't need that.
profile.set_preference("browser.urlbar.autoFill", False) # Disable autofill on URL bar.
profile.set_preference("browser.urlbar.autocomplete.enabled", False) # Disable autocomplete on URL bar.
profile.set_preference("browser.urlbar.showPopup", False) # Disable list of URLs when typing on URL bar.
profile.set_preference("browser.urlbar.showSearch", False) # Disable search bar.
profile.set_preference("extensions.checkCompatibility", False) # Addon update disabled
profile.set_preference("extensions.checkUpdateSecurity", False)
profile.set_preference("extensions.update.autoUpdateEnabled", False)
profile.set_preference("extensions.update.enabled", False)
profile.set_preference("general.startup.browser", False)
profile.set_preference("plugin.default_plugin_disabled", False)
profile.set_preference("permissions.default.image", 2) # Image load disabled 

#Pick the city
city = 'Milwaukee, WI, United States'
url = 'https://www.airbnb.com/?_set_bev_on_new_domain=1595286208_s%2FfIxrqWUv%2BUDLLa&locale=en'
browser = webdriver.Firefox(firefox_profile=profile)



############################################################# Search for results and copy links
#Open URL
print('Acessing URL...')
browser.get(url)
print('Loading home page')

#Wait to load
time.sleep(5)

#print('Searching for acommodations in %s...'%location)

try:
    
    #Send keys to search location
    print('Typing location')
    search = WebDriverWait(browser,20).until(
    EC.visibility_of_element_located((By.ID, "bigsearch-query-detached-query"))).send_keys(city)

    #Submit
    print('Submiting...')
    browser.find_element_by_class_name("_m9v25n").click()

    #Wait
    print('Waiting new page to load')
    time.sleep(2)
    
except Exception as e:
    print("Failed to search for location")
    print(e)
    sys.exit() 

#########Pages

page = 0
links = []

#Copy the URLs
while True:

    #Count page number
    page += 1
    
    #Get all acomodation links   
    print('Page', page,': copying acommodation URLs...')

    items = WebDriverWait(browser,50).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "_gjfol0")))
    numLink = 0

    for item in items:
        try:
            numLink += 1
            link = item.get_attribute('href')
            links.append(link)
            print('Link %s coppied!'%numLink)
            link_copied = True
        except Exception as e:
            print('Not possible to copy link %s'%numLink)
            print(e)

    print('Page %s done!'%page)
        
    #Scroll down
    htmlElem = browser.find_element_by_tag_name('html')
    htmlElem.send_keys(Keys.END)
    htmlElem.send_keys(Keys.PAGE_UP)

    #Click next button    
    print('Going to the next page...')
    try:
        nextButton = WebDriverWait(browser,10).until(EC.element_to_be_clickable((By.XPATH, "//a[@aria-label = 'Next']")))
        nextButton.click()

        #Wait
        print("waiting new page to load")
        time.sleep(2)

        
    except Exception as ex:
        print('Last page.')
        break
print('End of reading links.')
######################################################################




##################################################### CREATE WORKBOOK
print('Creating new workbook...')

os.chdir('C:\\Users\\Gabriela Avila\\Desktop\\Automating the boring stuff')

wb = openpyxl.Workbook()
ws = wb.active

#Define columns names
columns = {1: 'Title',
           2: 'People',
           3: 'Bedrooms',
           4: 'Bathrooms',
           5: 'Reviews',
           6: 'Price',
           7: 'Amenities',
           8: 'Description',
           9: 'Link'}

for i in range(1, 10):
        ws.cell(row = 1, column = i).value = columns[i]
        ws.cell(row = 1, column = i).font = Font(bold = True)

################################################## ACCESS LINKS
lin = 0
for link in links:

    lin+=1
    print('Accessing link', lin)

    browser.get(link)
    #Wait
    print("waiting new page to load")
    time.sleep(2)
    
    #WAIT UNTIL FIND BUTTON, TO MAKE SURE PAGE IS FULL LOADED
    button = WebDriverWait(browser,50).until(EC.presence_of_element_located((By.CLASS_NAME, "_13e0raay")))
    r = browser.page_source
    soup = BeautifulSoup(r, 'html.parser')

    #Get TITLE
    title = soup.h1.text
    print('Title:',title)

    #Get people/bedrooms/bathrooms
    #Iterate over text and get the amount of people, bedrooms and bathrooms
    #text = [<span>2 hóspedes</span>, <span aria-hidden="true"> · </span>, <span>1 quarto</span>, <span aria-hidden="true"> · </span>, <span>1 cama</span>, <span aria-hidden="true"> · </span>, <span>1 banheiro</span>]

    t = soup.find('div', class_ = '_tqmy57')
    text = t.find_all('span')

    people = None
    bedrooms = None
    bathrooms = None

    for tag in range(len(text)):
        if tag == 0:
            people = (text[tag].get_text())
        elif tag == 2:
            bedrooms = (text[tag].get_text())
        elif tag == 6:
            bathrooms = (text[tag].get_text())
            
    print('People:', people)
    print('Bedrooms:', bedrooms)
    print('Bathrooms:', bathrooms)

    #Get price
    try:
        p = soup.find('span', class_ = '_pgfqnw')
        price = p.get_text()
    except Exception as e:
        price = 'None'
    print('Price:', price)
    
    #Get reviews
    try:
        review = soup.find("button", class_="_1wlymrds")["aria-label"]
    except Exception as e:
        review = 'None'
    print('Reviews:', review)
    #Get amenities
    #Find the button
    try:
        #button = browser.find_element_by_class_name('_13e0raay')

        #Scroll to button
        button.location_once_scrolled_into_view
        htmlElem = browser.find_element_by_tag_name('html')

        htmlElem.send_keys(Keys.UP)
        htmlElem.send_keys(Keys.UP)
        button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "_13e0raay")))
        button.click()

        #Find amenities
        sectionList = WebDriverWait(browser,10).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class = '_1cnse2m']")))

        amenities = ''''''
        n =0
        for item in sectionList:
            n +=1
            if n == 2:
                amenities = str(item.text)

        #Format text
        amenities = amenities.replace('\n','.')
        amenities = re.sub(r'(Amenities.|Basic.|Essentials.|Dining.|Facilities.|Bed and bath.|Logistics.|Outdoor.|Not included.*)',"", amenities)      

    except Exception as e:
        print(e)
        amenities = 'None'
    print('Amenities:', amenities)

    #Get description
    try:
        d = soup.find('div', class_ =  '_1y6fhhr')
        description = d.span.text
        
    except Exception as e:
        print(e)
        description = 'None'
    print('Description:', description)
    
    #WRITE CONTENT
    print('Writing in file...')
    ws.append((title, people, bedrooms, bathrooms, review, price, amenities, description, link))    

    wb.save('AirbnbTestNEW.xlsx')
    print('Done!!!!!!!!!!!!')

wb.save('AirbnbTest6.xlsx')
browser.quit()



















