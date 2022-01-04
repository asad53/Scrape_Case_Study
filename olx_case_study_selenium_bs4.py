#Libraries

import requests
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.touch_actions import TouchActions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from fake_useragent import UserAgent
import time
import openpyxl
import pandas
import datetime
from bs4 import BeautifulSoup as soup



def configure_driver():

    # Add additional Options to the webdriver
    chrome_options = Options()

    ua = UserAgent()
    userAgent = ua.random  # THIS IS FAKE AGENT IT WILL GIVE YOU NEW AGENT EVERYTIME
    print(userAgent)
    chrome_options.add_argument("--headless")                    #if you don't want to see the display on chrome just uncomment this
    chrome_options.add_argument(f'user-agent={userAgent}')  # useragent added
    chrome_options.add_argument("--log-level=3")  # removes error/warning/info messages displayed on the console
    chrome_options.add_argument("--disable-notifications")  # disable notifications
    chrome_options.add_argument("--disable-infobars")  # disable infobars ""Chrome is being controlled by automated test software"  Although is isn't supported by Chrome anymore
    chrome_options.add_argument("start-maximized")  # will maximize chrome screen
    chrome_options.add_argument('--disable-gpu')  # disable gpu (not load pictures fully)
    chrome_options.add_argument("--disable-extensions")  # will disable developer mode extensions

    prefs = {"profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)             #we have disabled pictures (so no time is wasted in loading them)

    #Configure the final driver with all options
    driver = webdriver.Chrome(ChromeDriverManager().install(),options=chrome_options)  # you don't have to download chromedriver it will be downloaded by itself and will be saved in cache

    #Return updated driver
    return driver



def setup_worksheet():

    print("")


    #Constructing Excel Name
    output_filename = "OLX_Data.xlsx"
    print("SAVING TO: ", output_filename)
    print("")
    print("")

    # Intializing OLX data Workbook
    wb_olx = openpyxl.Workbook()
    olx_sheet = wb_olx.active
    olx_sheet.title = "Sheet1"

    # Intializing the column names now
    header_row = ['listing_url', 'listing_date', 'listing_id', 'condition', 'type', 'description',
                  'price','seller_name', 'is_featured']
    olx_sheet.append(header_row)
    wb_olx.save(output_filename)

    #Returned made excel
    return wb_olx,olx_sheet,output_filename


def convert_to_csv():

    excel_data_df = pandas.read_excel("OLX_Data.xlsx",sheet_name="Sheet1")

    excel_data_df.to_csv("OLX_Data_Csv.csv",index=False)

    print("")
    print("Saved To Csv")
    print("")




def RunScrapper(driver,wb_olx,olx_sheet,output_filename):

    #Start Time Of Scrapper
    start_time = time.time()



    #Get To Link
    driver.get("https://www.olx.com.pk/tablets_c1455")


    #Loop For It To Wait And Click On Load More --  To Extract all listings urls
    #Used Selenium Here
    loop_status=False
    print("")
    print("Clicking All Load More")
    print("")
    pg_no=1
    while loop_status!=True:

        try:
            print("Page No: ",pg_no)
            pg_no+=1
            WebDriverWait(driver, 4).until(expected_conditions.element_to_be_clickable((By.XPATH, '//span[text()="Load more"]')))
            #Click the Load More Button
            driver.find_element_by_xpath('//span[text()="Load more"]').click()
        except Exception:
            print("")
            print("No Load More! Initiating Links Retrieval!")
            print("")
            break

    #defining Ad Links
    ad_links=[]

    #List of all the ads on page using beautiful soup
    pagesoup = soup(driver.page_source, "html.parser")
    container = pagesoup.findAll("li", {"aria-label": "Listing"})


    for contain in container:
        #Retrieve link
        link="https://www.olx.com.pk"+contain.find('a').get('href')

        #Append to Total List
        ad_links.append(link)


    #remove duplicate links
    ad_links = list(dict.fromkeys(ad_links))

    #Get total number of ad links
    print("")
    total_ads=len(ad_links)
    print("Total Ads To Iterate: ",total_ads)
    print("")

    #Initiate Ad Number
    ad_no=1


    #Iterate through each ad link to scrape details using beasutiful soup
    for ad_link in ad_links:
        print("Ad No: ",ad_no,"/",total_ads)
        ad_no+=1
        print("Ad Link: ",ad_link)

        try:
            # Parse using html Parser
            page = requests.get(ad_link)
            pagesoup = soup(page.text, 'html.parser')

            # Retrieved Seller Name And Cleaned it
            try:
                seller_name = pagesoup.find("div", {"class": "_1075545d _6caa7349 _42f36e3b d059c029"})
                seller_name = seller_name.find("span").text
                seller_name = seller_name.strip()
            except Exception:
                seller_name = ''
                pass

            # Retrieved Price And Cleaned it
            try:
                ad_price = pagesoup.find("span", {"class": "_56dab877"}).text

                try:
                    ad_price = ad_price.replace("Rs", "")
                except Exception:
                    pass

                try:
                    ad_price = ad_price.replace(",", "")
                except Exception:
                    pass

                ad_price = ad_price.strip()
            except Exception:
                ad_price = ''
                pass

            # Retrieved Ad Id And Cleaned it
            try:
                ad_Id = pagesoup.find("div", {"class": "_171225da"}).text

                try:
                    ad_Id = int(ad_Id.replace("Ad id ", "").strip())
                except Exception:
                    pass
            except Exception:
                ad_Id = 0
                pass

            # Retrieved Ad Description And Cleaned it
            try:
                ad_Description = pagesoup.find("div", {"class": "_0f86855a"}).text

                try:
                    ad_Description = ad_Description.encode('UTF-8')
                except Exception:
                    pass

                ad_Description = ad_Description.strip()

            except Exception:
                ad_Description = ''
                pass

            # Retrieved Ad Feature and Cleaned it
            try:
                ad_feature = pagesoup.find("span", {"class": "_8918c0a8 _2e82a662 a695f1e9']"}).text
                if ad_feature == 'Featured':
                    ad_feature = True
                else:
                    ad_feature = False
            except Exception:
                ad_feature = False
                pass

            # Retrieved Ad Date and Cleaned it and Formated it
            try:
                overview = pagesoup.find("div", {"aria-label": "Overview"})
                overview_container = overview.findAll("span", {"class": "_8918c0a8"})
                ad_Date = ''
                # We have multiple class names so used keyword ago to search out relevant one
                for overview_contain in overview_container:
                    try:
                        value_search = overview_contain.find("span").text
                        if "ago" in value_search:
                            ad_Date = value_search
                            break
                    except Exception:
                        pass
                if ad_Date != '':
                    value = ad_Date
                    # strip it first
                    value = value.strip()
                    # split using spaces
                    duration = value.split(" ")
                    # get the week,day,month,year,minute,hour,second variable
                    duration_heading = duration[1]

                    # check which one it matches to and perform actions accordingly using current date
                    if "week" in duration_heading:
                        tod = datetime.datetime.now()
                        d = datetime.timedelta(weeks=int(duration[0].strip()))
                        value = tod - d
                    elif "day" in duration_heading:
                        tod = datetime.datetime.now()
                        d = datetime.timedelta(days=int(duration[0].strip()))
                        value = tod - d
                    elif "minute" in duration_heading:
                        value = datetime.datetime.now()
                    elif "hour" in duration_heading:
                        value = datetime.datetime.now()
                    elif "second" in duration_heading:
                        value = datetime.datetime.now()
                    elif "month" in duration_heading:
                        tod = datetime.datetime.now()
                        d = datetime.timedelta(weeks=int(duration[0].strip()) * 4)
                        value = tod - d
                    else:
                        pass

                    # return date in end
                    ad_Date = value.date()
            except Exception:
                ad_Date = ''
                pass

            # Retrieved Ad Type and Ad Condition and cleaned it

            try:
                ad_Type = ''
                ad_Condition = ''
                details = pagesoup.findAll("div", {"class": "_676a547f"})
                for detail in details:
                    # get the heading of details and content written with it
                    # this is done because there is no specific classname id or aria label given for condition and type
                    spans = detail.findAll("span")
                    heading = spans[0].text
                    content = spans[1].text

                    # match condition
                    if heading == "Condition":
                        ad_Type = content

                    # match type
                    elif heading == "Type":
                        ad_Condition = content
                    else:
                        pass
            except Exception:
                ad_Type = ''
                ad_Condition = ''
                pass

            # Print Values On Console

            print("Seller Name: ", seller_name)
            print("Ad Price: ", ad_price)
            print("Ad Id: ", ad_Id)
            print("Ad Feature: ", ad_feature)
            print("Ad Date: ", ad_Date)
            print("Ad Condition: ", ad_Condition)
            print("Ad Type: ", ad_Type)
            print("Ad Description: ", ad_Description)

            #Check if some value is there
            if ad_Date=='' and ad_Id=='' and ad_Condition=='' and ad_Type=='' and ad_Description=='' and ad_price=='' and ad_feature==False and seller_name=='':
                pass
            else:
                # Saved in Excel File
                value_row = [ad_link, ad_Date, ad_Id, ad_Condition, ad_Type, ad_Description,
                             ad_price, seller_name, ad_feature]
                olx_sheet.append(value_row)
                wb_olx.save(output_filename)

        except Exception:
            print("Error In Retrieving Url! Moving To Next One")
            pass

        #Moving to Next One

        print("  ")
        print("****************************************************************")
        print("  ")









    # give time taken to execute everything
    print("time elapsed: {:.2f}s".format(time.time() - start_time))



if __name__ == '__main__':

    # Setting Up Worksheet
    wb_olx,olx_sheet,output_filename=setup_worksheet()

    # create the driver object.
    driver = configure_driver()

    # call the scrapper to run
    RunScrapper(driver,wb_olx,olx_sheet,output_filename)


    #Converted xlsx to csv
    convert_to_csv()

    # close the driver after execution
    driver.close()
















