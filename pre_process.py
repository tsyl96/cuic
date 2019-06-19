
#!/usr/bin/python3
from selenium import webdriver
from datetime import date
import datetime
import time
import csv
import os
from pathlib import Path

today = date.today()
d_out = today.strftime("%d.%m.%y")
hr = datetime.datetime.now().hour
out_folder = Path('/home/thetsu/Desktop/autoprogram/cuic')
#const_time = '(' + str(5) + '-' + str(6) + ')'
#out_name = d_out + const_time + ".csv"
display_filename = "Hourly Report for(" + d_out + ").csv"
with open(str( out_folder / display_filename), mode='w') as file:
    file_writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    file_writer.writerow(
        ['Interval', 'CallOffered', 'CallsHandled', 'AbanRate', 'SLPercentage', 'ASA', 'AHT', 'AvgTalkTime', 'AvgHoldTime', 'Short Call'])
