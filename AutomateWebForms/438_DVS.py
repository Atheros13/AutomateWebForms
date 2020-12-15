import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome(ChromeDriverManager().install())

driver.maximize_window()  
driver.get("https://www7.eaccounts.co.nz/eLogin_Main.asp")  
time.sleep(2)
driver.find_element_by_name("User__Name").send_keys("TCC16")
time.sleep(2)
driver.find_element_by_name("User__Pass").send_keys("dvs08")
time.sleep(2)
driver.find_element_by_name('Login').click()
time.sleep(0.9)
driver.find_element_by_class_name("MENU-BUTTON").click()
time.sleep(2)
driver.find_element_by_name("Load_Prospect").click()

time.sleep(2)
driver.find_element_by_name("Prospect_Loaded_By").send_keys("Michael Atheros")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Contact").send_keys("The Call Centre")

time.sleep(0.9)
driver.find_element_by_name("Prospect_Name").send_keys("The Call Centre")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Del_Add_1").send_keys("4 Daly Street")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Del_Add_2").send_keys("Lower Hutt")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Del_Add_3").send_keys("Wellington")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Del_Add_4").send_keys("")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Del_Zip").send_keys("5011")

time.sleep(0.9)
driver.find_element_by_name("Prospect_Suburb").send_keys("Hutt Central / Hutt Valley = Matt Richardson")

time.sleep(0.9)
driver.find_element_by_name("Prospect_Ph").send_keys("0800 000 000")
time.sleep(0.9)
driver.find_element_by_name("Prospect_CellPh").send_keys("022 000 000")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Email").send_keys("none@none.com")
 
time.sleep(0.9)
driver.find_element_by_name("Prospect_Source").send_keys("[3] Website Email Lead")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Note_Type").send_keys("Call Centre: Sales Lead")

time.sleep(0.9)
driver.find_element_by_name("Prospect_Wanted").send_keys("DVS Consultation")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Wanted2").send_keys("")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Wanted3").send_keys("")

time.sleep(0.9)
driver.find_element_by_name("Prospect_Relevant").send_keys("")
time.sleep(0.9)
driver.find_element_by_name("Prospect_Best_Time").send_keys("")

time.sleep(0.9)
#driver.find_element_by_name("Save_Prospect").click()