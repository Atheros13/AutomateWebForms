import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import shutil

# open and import all Wilsons Coupon Data
import xlrd
sheet = xlrd.open_workbook("CouponAutomation.xls").sheet_by_index(0)

# open coupon site and log in
driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get("http://request.couponcompany.co.nz/service-request.aspx")  
driver.find_element_by_id("txtLoginUsername").send_keys("support@wilson.co.nz")
driver.find_element_by_id("txtLoginPassword").send_keys("Wil$ons")
time.sleep(0.9)
driver.find_element_by_class_name('btn').click()
time.sleep(0.9)

# coupon count
coupon_count = 0

# run loop for automating
for r in range(1, sheet.nrows-2):

    # check if all required columns have data, ignore if they do not
    check = True
    for c in range(8):
        if sheet.cell_value(rowx=r, colx=c) == "":
            check = False
    if not check:
        print(True)
        continue

    # assign column value to variables
    name = sheet.cell_value(rowx=r, colx=0).split()
    if len(name) == 2:
        firstname = name[0]
        lastname = name[1]
    elif len(name) > 2:
        firstname = name[0]
        lastname = " ".join(name[1:])
    else:
        firstname = name[0]
        lastname = name[0]

    email = sheet.cell_value(rowx=r, colx=2)
    description = sheet.cell_value(rowx=r, colx=3)
    date = sheet.cell_value(rowx=r, colx=4)
    if '/' in date:
        date = date.split('/')
        if len(date[0]) == 1:
            date[0] = '0%s' % date[0]
        if len(date[1]) == 1:
            date[1] = '0%s' % date[1]
        date = "".join(date)
    elif '-' in date:
        d = date.split('-')
        date = '%s%s%s' % (d[2], d[1], d[0])
    else:
        continue

    subject = sheet.cell_value(rowx=r, colx=5)
    email_body = sheet.cell_value(rowx=r, colx=6)
    reference = sheet.cell_value(rowx=r, colx=7)

    driver.find_element_by_id("input1726").send_keys(firstname)
    driver.find_element_by_id("input1727").send_keys(lastname)
    driver.find_element_by_id("input1728").send_keys(email)
    driver.find_element_by_id("input1729").send_keys(description)
    expiry_date = driver.find_element_by_id("input1730")
    expiry_date.send_keys(date)
    #date.send_keys(Keys.TAB)
    #date.send_keys("2020")
    driver.find_element_by_id("input1733").send_keys(subject)
    driver.find_element_by_id("input1734").send_keys(email_body)
    driver.find_element_by_id("input1735").send_keys(reference)
    time.sleep(0.9)
    #driver.find_element_by_class_name("action-submit").click()
    time.sleep(0.9)    
    #driver.find_element_by_id("input1737").click()
    
    #
    coupon_count += 1

## SEND JOHN EMAIL

sys.exit(0)

import smtplib
import ssl

# User configuration
sender_email = 'thecallcentre4@gmail.com'
receiver_emails = ['m.t.atheros@gmail.com','john@thecallcentre.co.nz']
password = 'Naz55105'

# Email text
if coupon_count > 0:
    email_body = '\n\n847 WILSON COUPONS: x %s' % coupon_count
else:
    email_body = '\n\nThe Wilsons Coupon code was run, but there were no coupons in the report.'


# Creating a SMTP session | use 587 with TLS, 465 SSL and 25
server = smtplib.SMTP('smtp.gmail.com', 587)
# Encrypts the email
context = ssl.create_default_context()
server.starttls(context=context)
# We log in into our Google account
server.login(sender_email, password)
# Sending email from sender, to receiver with the email body
server.sendmail(sender_email, receiver_emails, email_body)
server.quit()

## CLEAR COUPON SHEET (TO PREVENT ERROR)
import os
if os.path.exists("G:\\Customer Reporting\847 - Wilson\\Coupon Automation\\CouponAutomation.xls"):
  os.remove("G:\\Customer Reporting\847 - Wilson\\Coupon Automation\\CouponAutomation.xls")
else:
  print("The file does not exist")

shutil.copy("G:\\Customer Reporting\847 - Wilson\\Coupon Automation\\TestData\\CouponAutomation.xls", 
            "G:\\Customer Reporting\847 - Wilson\\Coupon Automation\\CouponAutomation.xls")

## CLOSE 
driver.quit()