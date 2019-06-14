from selenium import webdriver
from datetime import date
import datetime
import time
import csv
import os

today = date.today()
d_out = today.strftime("%d.%m.%y")
hr = datetime.datetime.now().hour
#const_time = '(' + str(5) + '-' + str(6) + ')'
#out_name = d_out + const_time + ".csv"
display_filename = "Hourly Report for(" + d_out + ").csv"
with open(display_filename, mode='w') as file:
    file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    file_writer.writerow(
        ['Interval', 'CallOffered', 'CallsHandled', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime', 'AbanRate',
         'SLPercentage','Short Call'])