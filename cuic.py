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

chrome_options = webdriver.ChromeOptions();
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome("/usr/local/bin/chromedriver",chrome_options=chrome_options)

driver.get("https://10.84.86.44:8444/cuic/Main.htmx")


def cuic_login():
    driver.find_element_by_id("j_username").send_keys("C64610")
    driver.find_element_by_id("j_password").send_keys("Shine@321")
    driver.find_element_by_id("j_domain").send_keys("CUIC")
    driver.find_element_by_id("cuesLoginSubmitButton").click()


cuic_login()


### OSCC
###Stock Report
###CCSC Dashboard all

def open_CCSC():
    driver.find_element_by_id("reportDrawerToggler").click()
    driver.implicitly_wait(1)
    # driver.find_element_by_id("reportDrawer_0").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_9").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_22").click()


open_CCSC()

out_folder = Path('/home/thetsu/Desktop/autoprogram/cuic')

today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
print(d3)
hr = datetime.datetime.now().hour
time = '(' + str(hr - 1) + '-' + str(hr) + ')'

total_hr_time = '(' + str(hr-2) + '-' + str(hr-1) + ')'

print('hr time for total report=' + total_hr_time)

start_hr = str(hr - 1) + ':00'
end_hr = str(hr) + ':00'
insert_hr = start_hr + '-' + end_hr

print("for one hr = " + insert_hr)

out_name = d_out + time + ".csv"
excel_name = 'Hourly Report For ' + d_out + time + ".xlsx"
old_name = d_out + total_hr_time + ".csv"
display_filename = "Hourly Report for(" + d_out + ").csv"
total_name = "Total Report" + d_out + time + ".csv"
final_filename = "Final " + d_out + ".csv"

print("old name =" + old_name)

'''
today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
print(
hr = datetime.datetime.now().hour
out_name = d_out + "(" + str(hr - 1) + "-" + str(hr) + ").csv"
'''


def data_output():
    # driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
        driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))
    except NoSuchElementException as exception:
        pass
    # start date test by one day

    # start date
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-dateParamTextbox").send_keys(d3)

    # end date
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-dateParamTextbox").send_keys(d3)

    # start hour,min and second
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeMinParamTextbox").send_keys("00")
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeSecParamTextbox").send_keys("00")

    # end hour,min and second
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeMinParamTextbox").send_keys("59")
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeSecParamTextbox").send_keys("59")

    driver.find_element_by_id("RunFilterBtnlink").click()

    # switch_frame

    driver.implicitly_wait(5)

    # must to switch default content

    driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
        driver.switch_to.frame(driver.find_element_by_id('view1_iframe'))
        driver.switch_to.frame(driver.find_element_by_id('viewframe'))
    except NoSuchElementException as exception:
        pass


    # web_table = driver.find_element_by_class_name('dataTable')

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
    print(CallsOffered.text)
    print(CallsHandled.text)
    print(ASA.text)
    print(AHT.text)
    print(AvgTalkTime.text)
    print(AvgHoldTime.text)
    print(AbanRate.text)
    print(SLPercentage.text)

    '''
    with open('newfile.csv', mode= 'w') as csv_file:
        fieldnames = ['CallsOffered','CallsHandled','ASA','AHT','AvgTalkTime','AvgHoldTime','AbanRate','SLPercentage']
        writer = csv.DictWriter(csv_file,fieldnames = fieldnames)
        writer.writeheader()
        writer.writerows({'CallsOffered' : CallsOffered.text , 'CallsHandled' : CallsHandled.text , 'ASA' : ASA.text , 'AHT' : AHT.text , 'AvgTalkTime' : AvgTalkTime.text ,'AvgHoldTime' : AvgHoldTime.text , 'AbanRate' : AbanRate.text , 'SLPercentage' : SLPercentage.text })

    '''

    # out_name = d_out + "(" + str(hr-1) + "-" + str(hr) + ").csv"

    with open(str(out_folder / out_name), mode='w') as file:
        file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        file_writer.writerow(
            ['Interval','CallOffered', 'CallsHandled', 'AbanRate', 'SLPercentage', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime'])
        file_writer.writerow(
            [insert_hr ,CallsOffered.text, CallsHandled.text,
             AbanRate.text, SLPercentage.text + "%", ASA.text, AHT.text, AvgTalkTime.text, AvgHoldTime.text])

    print("for one report= " + out_name)


data_output()

driver.refresh()


def reopen_CCSC():
    driver.switch_to.default_content()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_19").click()


reopen_CCSC()


def short_call_data():
    # switch frame to date field

    driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
    driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))

    # fill date

    # start date
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-dateParamTextbox").send_keys(d3)

    # end date
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-dateParamTextbox").send_keys(d3)

    # start hour,min and second
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeMinParamTextbox").send_keys("00")
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeSecParamTextbox").send_keys("00")

    # end hour,min and second
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeMinParamTextbox").send_keys("59")
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeSecParamTextbox").send_keys("59")

    driver.find_element_by_id("RunFilterBtnlink").click()

    # switch_frame

    driver.implicitly_wait(5)

    # must to switch default content

    driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
        driver.switch_to.frame(driver.find_element_by_id("view1_iframe"))
        driver.switch_to.frame(driver.find_element_by_id("viewframe"))

    except NoSuchElementException as exception:
        pass

    #  driver.switch_to.frame(driver.find_element_by_id('viewFrameBody'))

    # get short call data
    shortcall_table = driver.find_element_by_id('dataTable')


    short_call = shortcall_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
    print(short_call.text)

    data = pd.read_csv(str(out_folder / out_name))
    # data = data.convert_objects(convert_numeric=True)
    data.insert(9,'Short Call', short_call.text)
    data.to_csv(str(out_folder / out_name),index=False)

short_call_data()

userDf = pd.read_csv(str(out_folder / out_name))
print(userDf)

with open(str(out_folder / display_filename), 'a')as append_total:
    userDf.to_csv(append_total,index=False ,header=False)

#with ExcelWriter(excel_name)as ew:
#    pd.read_csv(display_filename).to_excel(ew, sheet_name=display_filename)


#get data for total

driver.forward()
driver.back()


def total_CCSC():
    #driver.switch_to.default_content()
    #driver.implicitly_wait(3)
    driver.find_element_by_id("reportDrawerToggler").click()
    driver.implicitly_wait(1)
    # driver.find_element_by_id("reportDrawer_0").click()
    driver.implicitly_wait(1)
#    driver.find_element_by_id("reportDrawer_9").click()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_22").click()

total_CCSC()

def total_output():
    # driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
        driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))
    except NoSuchElementException as exception:
        pass
    # start date test by one day

    # start date
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-dateParamTextbox").send_keys(d3)

    # end date
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-dateParamTextbox").send_keys(d3)

    # start hour,min and second
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeHourParamTextbox").send_keys("6")
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeMinParamTextbox").send_keys("00")
    driver.find_element_by_id("A7E72A86100001610016C8520A54562C-timeSecParamTextbox").send_keys("00")
    driver.implicitly_wait(3)

    # end hour,min and second
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeMinParamTextbox").send_keys("59")
    driver.find_element_by_id("A7E72A86100001610016C8530A54562C-timeSecParamTextbox").send_keys("59")
    driver.implicitly_wait(3)

    driver.find_element_by_id("RunFilterBtnlink").click()

    # switch_frame

    driver.implicitly_wait(5)

    # must to switch default content

    driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
        driver.switch_to.frame(driver.find_element_by_id('view1_iframe'))
        driver.switch_to.frame(driver.find_element_by_id('viewframe'))
    except NoSuchElementException as exception:
        pass


    # web_table = driver.find_element_by_class_name('dataTable')

    # get data from table
    web_table = driver.find_element_by_id('dataTable')
    total_CallsOffered = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[1]')
    total_CallsHandled = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[2]')
    total_ASA = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[4]')
    total_AHT = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[5]')
    total_AvgTalkTime = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[6]')
    total_AvgHoldTime = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[7]')
    total_AbanRate = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
    total_SLPercentage = web_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[10]')

    print("total calls offered", total_CallsOffered.text)
    print("total call handled", total_CallsHandled.text)
    print("total ASA", total_ASA.text)
    print("total AHT", total_AHT.text)
    print("total AvgTalkTime", total_AvgTalkTime.text)
    print("total AvgHoldTime", total_AvgHoldTime.text)
    print("total AbanRate", total_AbanRate.text)
    print("total SLPercentage", total_SLPercentage.text)

    with open(str(out_folder / total_name), mode='w') as file:
        file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        file_writer.writerow(
            ['Interval','CallOffered', 'CallsHandled', 'AbanRate', 'SLPercentage', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime'])
#        i = hr
#        while (i < 24):
#            final_interval = str(i) + ":00-" + str(i + 1) + ":00"
#            file_writer.writerow([final_interval,"", "", " ", "", "", "", "", "",""])
#            i += 1
        file_writer.writerow(
            ["Total", total_CallsOffered.text, total_CallsHandled.text, total_AbanRate.text, total_SLPercentage.text + "%",
             total_ASA.text, total_AHT.text, total_AvgTalkTime.text, total_AvgHoldTime.text])

    print("total report= " + total_name)
####Total Short Call

    driver.switch_to.default_content()
    driver.implicitly_wait(1)
    driver.find_element_by_id("reportDrawer_19").click()
    driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
    driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))

    # fill date

    # start date
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-dateParamTextbox").send_keys(d3)

    # end date
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-dateParamTextbox").send_keys(d3)

    # start hour,min and second
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeHourParamTextbox").send_keys(6)
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeMinParamTextbox").send_keys("00")
    driver.find_element_by_id("D6C74F8410000161001E9A060A54562C-timeSecParamTextbox").send_keys("00")

    # end hour,min and second
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeHourParamTextbox").send_keys(hr - 1)
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeMinParamTextbox").send_keys("59")
    driver.find_element_by_id("D6C74F8410000161001E9A070A54562C-timeSecParamTextbox").send_keys("59")

    driver.find_element_by_id("RunFilterBtnlink").click()

    # switch_frame

    driver.implicitly_wait(5)

    # must to switch default content

    driver.switch_to.default_content()
    try:
        driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
        driver.switch_to.frame(driver.find_element_by_id("view1_iframe"))
        driver.switch_to.frame(driver.find_element_by_id("viewframe"))

    except NoSuchElementException as exception:
       pass

    #  driver.switch_to.frame(driver.find_element_by_id('viewFrameBody'))

    # get short call data
    shortcall_table = driver.find_element_by_id('dataTable')
    total_short_call = shortcall_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
    print("Total ShortCall", total_short_call.text)

### Data Out

#shine
    data = pd.read_csv(str(out_folder / total_name))
    data.insert(9, 'Short Call', total_short_call.text)
    data.to_csv(str(out_folder / total_name), index=False)

#shine
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
    print(total_Df)

    shutil.copyfile(str(out_folder / display_filename), str(out_folder / final_filename))

    interval_df = pd.read_csv(str(out_folder / "Interval.csv"), error_bad_lines=False)
    print(interval_df)

    with open(str(out_folder / final_filename), 'a')as append_interval:
        interval_df.to_csv(append_interval, index=False, header=False)

    with open(str(out_folder / final_filename), 'a')as append_final:
        total_Df.to_csv(append_final, index=False, header=False)

    with ExcelWriter(str(out_folder / excel_name))as ew:
        pd.read_csv(str(out_folder / final_filename), error_bad_lines=False).to_excel(ew, sheet_name="MIS report", index=False)

total_output()

'''
#shine
    f = open(final_filename, 'r')
    reader = csv.reader(f)
    mylist = list(reader)
    print(mylist)
    f.close()
    mylist[19][9] = total_short_call.text
    my_new_list = open(final_filename, 'w', newline='')
    csv_writer = csv.writer(my_new_list)
    csv_writer.writerows(mylist)
    print(my_new_list)
    my_new_list.close()
    # shine
'''
    #shine
'''
   
'''
'''

    #append short call to csv
    with open(out_name, mode= 'a') as short_file:
        writer = csv.writer(short_file)



result = []

with open(out_name , mode= 'r')as f_read:
    fdata_read = csv.reader(f_read)
    for row in fdata_read:
        result.append(row)

with open(total_runfilename, mode="w+")as f_write:
    fdata_write = csv.writer(f_write)
    fdata_write.writerow

'''

'''
#table_id = driver.find_element_by_id('dataTable')
    #body = web_table.driver.find_element_by_id('tbody')
    #rows = body.find_element_by_tag_name('tr')
    #for row in rows:
        #Get the columns
     #   col = row.find_element_by_tag_name('td')[3]
        #print(col.text)

    #driver.find_element_by_xpath("//*[@id="tbody"]/table/tbody/")
   # driver.find_element_by_xpath("//*[@id='A7E72A86100001610016C8520A54562C-dateParamTextbox']").send_keys("05/23/2019")
    #self.driver.implicitly_wait(10)


#driver.delete_all_cookies()

//*[@id="tbody"]/tr/td[5]
//*[@id="tbody"]
//*[@id="dataTable"]#
//*[@id="scrollTableContainer"]
//*[@id="viewFrameBody"]
//*[@id="viewframe"]
//*[@id="gridhtml"]
//*[@id="view1_iframe"]
//*[@id="reportViewerDiv"]
//*[@id="cuesContentArea"]
/html/body
/html
//*[@id="RnR_CCSC_Dashboard_All"]
'''
