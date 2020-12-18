import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.keys import Keys

from outlook_msg import Message
from os import listdir
from os.path import isfile, join

import datetime
import sys
import shutil
import string
import random

## 
if len(sys.argv) > 1:

    argv = sys.argv

    loaded_by = "%s %s" % (argv[1], argv[2])
    username = argv[3]
    password = argv[4]

else:
    
    loaded_by = "Michael Atheros"
    username = "TCC16"
    password = "dvs08"

## FUNCTIONS

def is_integer(n):
    try:
        float(n)
    except ValueError:
        return False
    else:
        return float(n).is_integer()

def random_filename():
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(9))
    return result_str

## SET DEFAULT "" FOR VARIABLES

contact = ""
add1 = ""
add2 = ""
add3 = ""
add4 = ""
zip = ""

phone = ""
email = ""

wanted1 = "DVS Consultation"
wanted2 = ""
wanted3 = ""

## OPEN TEXT FILE AND EXTRACT DATA

files = [f for f in listdir("G:\\Customer Reporting\\438-DVS\\Sales Leads Automation\\") if isfile("G:\\Customer Reporting\\438-DVS\\Sales Leads Automation\\%s" % f)]

if len(files) == 0:

    sys.exit(0)

with open("G:\\Customer Reporting\\438-DVS\\Sales Leads Automation\\%s" % files[0]) as msg_file:
    msg = Message(msg_file)
    msg_file.close()

contents = msg.body
lines = contents.split('\n')

headings = []
for i in range(len(lines)):

    line = lines[i]
    
    if "Name" == line[:4] or "Name:" == line[:5]:
        contact = lines[i+1].strip()
    elif "Email" == line[:5] or "Email:" == line[:6]:
        email = lines[i+1].strip()
        try:
            email = email.split('<')[0]
        except:
            pass
    elif "Phone Number" == line[:12] or "Phone Number:" == line[:13] or "Phone" == line[:5] or "Phone:" == line[:6]:
        phone = lines[i+1].strip()
    elif "Address" == line[:7] or "Address:" == line[:7]:

        if "Book A Consultation" in files[0]:
            for a in range(1,5):
                txt = lines[i+a].strip()
                print(txt)
                if "Map It<html" == txt[:11]:
                    break

                if a > 1:
                    for words in txt.split():
                        if is_integer(words):
                            zip = words
                            txt = txt.split(zip)[0]
                            break
                if a == 1:
                    add1 = txt
                elif a == 2:
                    add2 = txt
                elif a == 3:
                    add3 = txt
                else:
                    add4 = txt

        elif "Free Consultation" in files[0]:
            address = lines[i+1].strip().split(',')
            for a in range(len(address)):
                txt = address[a]
                if a > 1:
                    for words in txt.split():
                        if is_integer(words):
                            zip = words
                            txt = txt.split(zip)[0]
                            break
                if a == 0:
                    add1 = txt
                elif a == 1:
                    add2 = txt
                elif a == 2:
                    add3 = txt
                else:
                    add4 = txt

    elif "Your Message (Optional)" in line and "Book A Consultation" in files[0]:

        message = lines[i+1].strip()
        if message != "":
            if len(message) > 70:
                wanted1 = message[:70]
                wanted2 = message[70:]
                if len(message) > 140:
                    wanted2 = message[70:140]
                    wanted3 = message[140:]

#print(contact, email, phone, add1, add2, add3, add4, zip)
#sys.exit(0)

## OPEN DRIVER
driver = webdriver.Chrome(ChromeDriverManager().install())

## LOGIN TO DVS AND GO TO LOAD PROSPECT
driver.get("https://www7.eaccounts.co.nz/eLogin_Main.asp")  
driver.find_element_by_name("User__Name").send_keys(username)
driver.find_element_by_name("User__Pass").send_keys(password)
time.sleep(0.9)
driver.find_element_by_name('Login').click()
time.sleep(0.9)
driver.find_element_by_class_name("MENU-BUTTON").click()
driver.find_element_by_name("Load_Prospect").click()

## PROCESS ALL DATA
driver.find_element_by_name("Prospect_Loaded_By").send_keys(loaded_by)
driver.find_element_by_name("Prospect_Contact").send_keys(contact)
driver.find_element_by_name("Prospect_Name").send_keys(contact)
driver.find_element_by_name("Prospect_Del_Add_1").send_keys(add1)
driver.find_element_by_name("Prospect_Del_Add_2").send_keys(add2)
driver.find_element_by_name("Prospect_Del_Add_3").send_keys(add3)
driver.find_element_by_name("Prospect_Del_Add_4").send_keys(add4)
driver.find_element_by_name("Prospect_Del_Zip").send_keys(zip)
driver.find_element_by_name("Prospect_Ph").send_keys(phone)
driver.find_element_by_name("Prospect_CellPh").send_keys("")
driver.find_element_by_name("Prospect_Email").send_keys(email)
driver.find_element_by_name("Prospect_Source").send_keys("[3] Website Email Lead")
driver.find_element_by_name("Prospect_Note_Type").send_keys("Call Centre: Sales Lead")
driver.find_element_by_name("Prospect_Wanted").send_keys(wanted1)
driver.find_element_by_name("Prospect_Wanted2").send_keys(wanted2)
driver.find_element_by_name("Prospect_Wanted3").send_keys(wanted3)
driver.find_element_by_name("Prospect_Relevant").send_keys("")
driver.find_element_by_name("Prospect_Best_Time").send_keys("")

## SEARCH CHECK
driver.find_element_by_name("dCust__Code").send_keys(add1)
driver.find_element_by_name("dCust__Code").send_keys(Keys.ENTER)


## HOLD WINDOW
time.sleep(3600)

## ON CLOSE

## MOVE FILE
del msg
shutil.move("G:\\Customer Reporting\\438-DVS\\Sales Leads Automation\\%s" % files[0], 
            "G:\\Customer Reporting\\438-DVS\\Sales Leads Automation\\Actioned\\%s%s%s" % (random_filename(), loaded_by, files[0]))

