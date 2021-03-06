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
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome("/usr/local/bin/chromedriver",chrome_options=chrome_options)

#open cuic website 
driver.get("https://10.84.86.44:8444/cuic/Main.htmx") 

# Extract time from date library
today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
hr = datetime.datetime.now().hour
time = '(' + str(hr - 1) + '-' + str(hr) + ')'
total_hr_time = '(' + str(hr-2) + '-' + str(hr-1) + ')'
start_hr = str(hr - 1) + ':00'
end_hr = str(hr) + ':00'
insert_hr = start_hr + '-' + end_hr
print (d3)
print (time)

# Define output file name and path
out_folder = Path('/home/thetsu/Desktop/autoprogram/cuic')
out_name = d_out + time + ".csv"
excel_name = 'Hourly Report For ' + d_out + time + ".xlsx"
old_name = d_out + total_hr_time + ".csv"
Header_file = "Hourly Report for(" + d_out + ").csv"
total_name = "Total Report" + d_out + time + ".csv"
final_filename = "Final " + d_out + ".csv"

# Login CUIC
def cuic_login():
    driver.find_element_by_id("j_username").send_keys("")
    driver.find_element_by_id("j_password").send_keys("")
    driver.find_element_by_id("j_domain").send_keys("CUIC")
    driver.find_element_by_id("cuesLoginSubmitButton").click()

cuic_login()

# Open CCSC dashboard
def open_CCSC():
    driver.find_element_by_id("reportDrawerToggler").click()
    driver.implicitly_wait(1)
    # driver.find_element_by_id("reportDrawer_0").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_9").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_23").click()

open_CCSC()

# Create header file at the morning
def Header_file_creation():
    with open(str( out_folder / Header_file), mode='w') as file:
        file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        file_writer.writerow(
            ['Interval', 'CallOffered', 'CallsHandled', 'AbanRate', 'SLPercentage', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime', 'Short Call'])
    print("\nHeader_file has been created")
        
if hr == 7:
    Header_file_creation()

# Class for CCSC Dashboard All
class MIS_Report:
    def __init__(self, report):
        self.report = report
        print (report)
#choose total or hourly
        if report == "totalreport":
            self.start_time = "6"
            self.file_output = total_name
            self.Interval_out = "Total"
            print("Total has been chose")
        elif report == "hourlyreport":
            self.start_time = str(hr-1)
            self.file_output = out_name
            self.Interval_out = insert_hr
            print("Hourly has been chose")
    
#Function for CCSC and short call output automation
    def data_output(self):
        print(self.start_time)
        print(self.file_output)
        print(self.Interval_out)
        try:
            driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
            driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))
        except NoSuchElementException as exception:
            pass
        

        #Enter start date
        driver.find_element_by_id("A7E72A86100001610016C8520A54562C-dateParamTextbox").send_keys(d3)

        #Enter end date
        driver.find_element_by_id("A7E72A86100001610016C8530A54562C-dateParamTextbox").send_keys(d3)

        #Enter start hour,min and second
        driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeHourParamTextbox").send_keys(self.start_time)
        driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeMinParamTextbox").send_keys("00")
        driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeSecParamTextbox").send_keys("00")

        #Enter end hour,min and second
        driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeHourParamTextbox").send_keys(hr - 1)
        driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeMinParamTextbox").send_keys("59")
        driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeSecParamTextbox").send_keys("59")

        driver.find_element_by_id("RunFilterBtnlink").click()

        driver.implicitly_wait(5)

        # switch to default content

        driver.switch_to.default_content()
        # switch_frame
        try:
            driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
            driver.switch_to.frame(driver.find_element_by_id('view1_iframe'))
            driver.switch_to.frame(driver.find_element_by_id('viewframe'))
        except NoSuchElementException as exception:
            pass

        # get data from table
        web_table = driver.find_element_by_id('dataTable')
        CallsOffered = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[1]')
        CallsHandled = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[2]')
        ASA = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[4]')
        AHT = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[5]')
        AvgTalkTime = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[6]')
        AvgHoldTime = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[7]')
        AbanRate = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
        SLPercentage = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[10]')
        
        # Write data in output file 
        with open(str(out_folder / self.file_output), mode='w') as file:
            file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            file_writer.writerow(
                ['Interval','CallOffered', 'CallsHandled', 'AbanRate', 'SLPercentage', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime'])
            file_writer.writerow(
                [self.Interval_out ,CallsOffered.text, CallsHandled.text,
                 AbanRate.text, SLPercentage.text + "%", ASA.text, AHT.text, AvgTalkTime.text, AvgHoldTime.text])

        print("\nAutomation done for " + self.file_output)
        
#### Short Call

        driver.switch_to.default_content()
        driver.implicitly_wait(1)
        driver.find_element_by_id("reportDrawer_20").click()
        driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
        driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))

        # Enter start date
        driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-dateParamTextbox").send_keys(d3)

        # Enter end date
        driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-dateParamTextbox").send_keys(d3)

        # Enter start hour,min and second
        driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeHourParamTextbox").send_keys(self.start_time)
        driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeMinParamTextbox").send_keys("00")
        driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeSecParamTextbox").send_keys("00")

        # Enter end hour,min and second
        driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeHourParamTextbox").send_keys(hr - 1)
        driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeMinParamTextbox").send_keys("59")
        driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeSecParamTextbox").send_keys("59")

        driver.find_element_by_id("RunFilterBtnlink").click()

        driver.implicitly_wait(5)

        # switch to default content

        driver.switch_to.default_content()
        # switch frame
        try:
            driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
            driver.switch_to.frame(driver.find_element_by_id("view1_iframe"))
            driver.switch_to.frame(driver.find_element_by_id("viewframe"))

        except NoSuchElementException as exception:
           pass

        # get short call data
        shortcall_table = driver.find_element_by_id('dataTable')
        total_short_call = shortcall_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
        #print("Total ShortCall", total_short_call.text)

        #Write short call data in output file
        data = pd.read_csv(str(out_folder / self.file_output))
        data.insert(9, 'Short Call', total_short_call.text)
        data.to_csv(str(out_folder / self.file_output), index=False)
        print("\nShortCall has been add in " + self.file_output)

# Call MIS Report Class for HourlyReport
HourlyReport = MIS_Report("hourlyreport")
HourlyReport.data_output()

# Hourly Output datas add to header file
def add_header():
    userDf = pd.read_csv(str(out_folder / out_name))
    #print(userDf)

    with open(str(out_folder / Header_file), 'a')as append_total:
        userDf.to_csv(append_total,index=False ,header=False)

    print("\nReport has been added to Header_file")

add_header()
driver.refresh()

#Open CCSC for TotalReport
def total_CCSC():
    #driver.switch_to.default_content()
    #driver.implicitly_wait(3)
    driver.find_element_by_id("reportDrawerToggler").click()
    driver.implicitly_wait(1)
    # driver.find_element_by_id("reportDrawer_0").click()
    driver.implicitly_wait(1)
#    driver.find_element_by_id("reportDrawer_9").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_23").click()

driver.forward()
driver.back()
total_CCSC()

#Call MIS Report For Total
TotalReport = MIS_Report("totalreport")
TotalReport.data_output()

#Combine Hourly and Total Reports
def output_combine():

    with open(str(out_folder / "Interval.csv"), mode='w') as file:
        file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        file_writer.writerow(
            ['Interval', 'CallOffered', 'CallsHandled', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime', 'AbanRate',
            'SLPercentage', 'Short Call'])
        i = hr
        while (i < 24):
            final_interval = str(i) + ":00-" + str(i + 1) + ":00"
            file_writer.writerow([final_interval, "", "", " ", "", "", "", "", "", ""])
            i += 1

    total_Df = pd.read_csv(str(out_folder / total_name))
    #print(total_Df)

    shutil.copyfile(str(out_folder / Header_file), str(out_folder / final_filename))

    interval_df = pd.read_csv(str(out_folder / "Interval.csv"), error_bad_lines=False)
    #print(interval_df)

    with open(str(out_folder / final_filename), 'a')as append_interval:
        interval_df.to_csv(append_interval, index=False, header=False)

    with open(str(out_folder / final_filename), 'a')as append_final:
        total_Df.to_csv(append_final, index=False, header=False)

    with ExcelWriter(str(out_folder / excel_name))as ew:
        pd.read_csv(str(out_folder / final_filename), error_bad_lines=False).to_excel(ew, sheet_name="MIS report", index=False)
    print("\nAutomation done for MIS Report" + excel_name)

output_combine()


#Send Mail
def automail():
	driver.implicitly_wait(1)
	yag=yagmail.SMTP('sbautomate19@gmail.com','ckvsrxnyjzsjbhco')

	SB=[
        	'koko.min@solutionbasket.com',
        	'shineaungthein@solutionbasket.com',
        	'thetsuyadanarlinn@solutionbasket.com',
		'shineaungthein@gmail.com'
	]
	
	yag.send(
        	to='rnr.it@telenor.com.mm',
        	cc=SB,
        	subject="Automate mail Hourly Report for RNR " + str(d_out),
        	contents="Dear All, \n \n Please Find hourly report for" + str(time) + ". \n\n Best Regards, \n Selenium",
        	attachments=str(out_folder / excel_name)
	)
	print("\nMail sent \n_______________________________________________________________________________________ \n")

automail()
