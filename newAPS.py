#!/usr/bin/python3
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date
from pathlib import Path
import datetime
import time
import csv
import pandas as pd
import os
import shutil
from selenium.common.exceptions import NoSuchElementException
from pandas.io.excel import ExcelWriter
from selenium.common.exceptions import StaleElementReferenceException
import yagmail

#open chrome with option
chrome_options = webdriver.ChromeOptions();
#chrome_options.add_argument('--kiosk')
#chrome_options.add_argument('window-size= 1980, 3000')
#chrome_options.add_argument('--headless')
#chrome_options.add_argument('--no-sandbox')
#chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome("/usr/local/bin/chromedriver",options=chrome_options)
driver.fullscreen_window()
#driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

#open cuic website
driver.get("https://10.84.86.44:8444/cuic/Main.htmx")


# Extract time from date library
today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
hr = datetime.datetime.now().hour


#file create
APS_file = "APS count for(" + d_out + ")"+ "-" + str(hr) + ".csv"

#folder create
# Define output file name and path
#out_folder = Path('/home/sbadmin/Desktop/autoprogram/cuic')
out_folder = Path('/Users/thetsuyadanarlinn/Desktop/ftthcuic/')

# Login CUIC
def cuic_login():
    driver.find_element_by_id("j_username").send_keys("C64610")
    driver.find_element_by_id("j_password").send_keys("Shine@321")
    driver.find_element_by_id("j_domain").send_keys("CUIC")
    driver.find_element_by_id("cuesLoginSubmitButton").click()

cuic_login()

# Open CCSC dashboard
def open_APS():
    driver.find_element_by_id("reportDrawerToggler").click()
    driver.implicitly_wait(1)
    # driver.find_element_by_id("reportDrawer_0").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_9").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_18").click()

open_APS()

def APS():
    driver.switch_to.frame(driver.find_element_by_id('RnR_Agent_Real_Time'))
    driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))
    driver.find_element_by_id("DA886B8810000161001F1B860A54562C-transferboxrightallimg").click()
    driver.find_element_by_id("RunFilterBtnlink").click()
    driver.implicitly_wait(5)
    #try:
    driver.switch_to.default_content()
    driver.switch_to.frame('RnR_Agent_Real_Time')
    driver.switch_to.frame('view1_iframe')
    driver.switch_to.frame('viewframe')
    #except NoSuchElementException as exception:
        #  pass
    #web_table = driver.find_element_by_id('dataTable')
    #CallsOffered = web_table.find_element_by_xpath('//*[@id="tbody"]/tr[3]/td')
    #ART_table = driver.find_element_by_id('dataTable')
    row_count = driver.find_elements_by_xpath('//*[@id="tbody"]/tr')
    #print(CallsOffered)
    count = str(len(row_count))
    print(count)
    print(len(row_count))

    with open(str(out_folder / APS_file) ,"w")as file:
        file.write(count)
        #file_writer = csv.writer(w)
        #file_writer.write(count)
APS()
