from selenium import webdriver
from datetime import date
import datetime
import time
import csv
import pandas as pd
import os

driver = webdriver.Chrome()

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

today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
print(d3)
hr = datetime.datetime.now().hour
time = '(' + str(hr - 1) + '-' + str(hr) + ')'
# print( "interval" + interval)
start_hr = str(hr - 1) + ':00'
end_hr = str(hr) + ':00'
insert_hr = start_hr + '-' + end_hr

# print("for one hr" + insert_hr)
out_name = d_out + time + ".csv"

'''
today = date.today()
d3 = today.strftime("%m/%d/20%y")
d_out = today.strftime("%d.%m.%y")
print(d3)
hr = datetime.datetime.now().hour
out_name = d_out + "(" + str(hr - 1) + "-" + str(hr) + ").csv"
'''


def data_output():
    # driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
    driver.switch_to.frame(driver.find_element_by_id('filter_iframe'))

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

    driver.switch_to.frame(driver.find_element_by_id('RnR_CCSC_Dashboard_All'))
    driver.switch_to.frame(driver.find_element_by_id('view1_iframe'))
    driver.switch_to.frame(driver.find_element_by_id('viewframe'))
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

    with open(out_name, mode='w') as file:
        file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        file_writer.writerow(
            ['CallOffered', 'CallsHandled', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime', 'AbanRate', 'SLPercentage'])
        file_writer.writerow(
            [CallsOffered.text, CallsHandled.text, ASA.text, AHT.text, AvgTalkTime.text, AvgHoldTime.text,
             AbanRate.text, SLPercentage.text])

    print(out_name)


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
    driver.switch_to.frame(driver.find_element_by_id('RnR_Answer_Short_Calls'))
    driver.switch_to.frame(driver.find_element_by_id('view1_iframe'))
    driver.switch_to.frame(driver.find_element_by_id('viewframe'))
    #  driver.switch_to.frame(driver.find_element_by_id('viewFrameBody'))

    # get short call data
    shortcall_table = driver.find_element_by_id('dataTable')
    short_call = shortcall_table.find_element_by_xpath('//*[@id="tbody"]/tr/td[8]')
    print(short_call.text)

    data = pd.read_csv(out_name)
    # data = data.convert_objects(convert_numeric=True)
    data.insert(8, 'Short call', short_call.text)
    data.to_csv(out_name)


short_call_data()

'''

    data = pd.read_csv(out_name)
    top = data.head()
    for row in top:
        p

    #append short call to csv
    with open(out_name, mode= 'a') as short_file:
        writer = csv.writer(short_file)

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