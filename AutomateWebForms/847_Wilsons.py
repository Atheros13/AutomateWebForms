import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome(ChromeDriverManager().install())

driver.maximize_window()  
driver.get("http://request.couponcompany.co.nz/service-request.aspx")  
time.sleep(0.9)
driver.find_element_by_id("txtLoginUsername").send_keys("support@wilson.co.nz")
time.sleep(0.9)
driver.find_element_by_id("txtLoginPassword").send_keys("Wil$ons")
time.sleep(0.9)
driver.find_element_by_class_name('btn').click()
time.sleep(0.9)
driver.find_element_by_id("input1726").send_keys("Michael")
time.sleep(0.9)
driver.find_element_by_id("input1727").send_keys("Atheros")
time.sleep(0.9)
driver.find_element_by_id("input1728").send_keys("m.t.atheros@gmail.com")
time.sleep(0.9)
driver.find_element_by_id("input1729").send_keys("Description Here")
time.sleep(0.9)
date = driver.find_element_by_id("input1730")
date.send_keys("1612")
date.send_keys(Keys.TAB)
date.send_keys("2020")
time.sleep(0.9)
driver.find_element_by_id("input1733").send_keys("Subject Here")
time.sleep(0.9)
driver.find_element_by_id("input1734").send_keys("Email Body Here")
time.sleep(0.9)
driver.find_element_by_id("input1735").send_keys("Reference Code")