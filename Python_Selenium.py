# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 19:18:56 2018

@author: j417062
"""

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui
import time
import pandas as pd
from bs4 import BeautifulSoup
import csv
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait    
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

user = "j417062"
pwd = "Jun"
pwd = pwd + "14"
pwd = pwd + "@93"

driver = webdriver.Chrome("C:\ProgramData\Anaconda3\Lib\site-packages\selenium\webdriver\common\chromedriver.exe")
driver.get("http://companyhost.com")
driver.maximize_window()
driver.find_element_by_id('username-id').send_keys(user)
driver.find_element_by_id('pwd-id').send_keys(pwd)
driver.find_element_by_name('login').click()

#driver.minimize_window()

delay = 3 # seconds
try:
    myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'reg_img_304316340')))
    print("Page is ready!")
except TimeoutException:
    print("Loading took too much time!")

driver.find_element_by_id('reg_img_304316340').click()
time.sleep(1)
driver.find_element_by_link_text('Incident Management').click()
time.sleep(1)
driver.find_element_by_link_text('Incident Management Console').click()
time.sleep(1)
driver.find_element_by_id('arid_WIN_6_303174300').click() #select drop down
time.sleep(1)
ActionChains(driver).key_down(Keys.DOWN).key_down(Keys.DOWN).key_down(Keys.DOWN).key_down(Keys.DOWN).key_down(Keys.DOWN).key_down(Keys.RETURN).perform()

#alrt = driver.switch_to_alert()
#alrt.accept()


###Breached tickets
driver.find_element_by_link_text('Breached').click()
table_id = driver.find_element_by_id('WIN_6_302087200').get_attribute('outerHTML') #all tickets
br  = pd.read_html(table_id)
print(br[1][0] + " - " + br[1][6] + " - " + br[1][2])
ZBreached = br[1]
###Unassigned
driver.find_element_by_link_text('Unassigned').click()
table_id = driver.find_element_by_id('WIN_6_302087200').get_attribute('outerHTML') #all tickets
br  = pd.read_html(table_id)
print(br[1][0] + " - " + br[1][6] + " - " + br[1][2])
ZUnassigned = br[1]
###Open
driver.find_element_by_link_text('Open').click()
table_id = driver.find_element_by_id('WIN_6_302087200').get_attribute('outerHTML') #all tickets
br  = pd.read_html(table_id)
print(br[1][0] + " - " + br[1][6] + " - " + br[1][2])
ZOpen = br[1]
#driver.find_element_by_class_name('TableSortImgDown').click()

########################### Unacknowledged queue monitoring

for x in range(0, 15):
    time.sleep(2)
    driver.find_element_by_link_text('Unassigned').click()
    driver.find_element_by_link_text('Unacknowledged').click()
    table_id = driver.find_element_by_id('WIN_6_302087200').get_attribute('outerHTML') #all tickets
    br  = pd.read_html(table_id)
    ZUnack = br[1]
    
    td_list = driver.find_elements_by_css_selector("#WIN_6_302087200 tr td") #------Incident table
    for index, row in ZUnack.iterrows():
        Inc_no = row[0]
        Prod = row[31]
        Exist_assignee = row[6]
        assignee = ""
        if Prod=='agrosoft':
            print("Agrosoft Ticket: " + Inc_no + " - " + Prod)
            assignee = "JOHN SMITH"
            #******** ASSIGNING THE TICKET TO RESOURSE
            for td in td_list:
                if(td.text == Inc_no):
                    ActionChains(driver).double_click(td).perform()
                    driver.find_element_by_xpath('//*[@id="WIN_7_1000000151"]/a/img').click()
                    myaction() # do it action
            ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
            driver.find_element_by_id("arid_WIN_7_1000000218").send_keys(assignee) #******** Set Assignee
            ActionChains(driver).key_down(Keys.DELETE).perform()
            ActionChains(driver).key_down(Keys.DOWN).perform()
            ActionChains(driver).key_down(Keys.RETURN).perform()
            driver.find_element_by_id("arid_WIN_7_7").click()
            driver.find_element_by_id("arid_WIN_7_7").send_keys("i") #******** Inprogress
            ActionChains(driver).key_down(Keys.RETURN).perform()
            driver.find_element_by_id("WIN_7_301614800").click() #********** Save
#            myaction() # do it action
            driver.find_element_by_link_text("Incident Console").click()
            driver.find_element_by_link_text('Unacknowledged').click()

########################## Action
def myaction():
    driverX = webdriver.Chrome("C:\ProgramData\Anaconda3\Lib\site-packages\selenium\webdriver\common\chromedriver.exe")
    time.sleep(1)
    driverX.get("http://fxtracker.test.cargill.com")
    driverX.maximize_window()
    time.sleep(1)
    driverX.get("http://fxtracker.test.cargill.com/user")
    time.sleep(1)
    driverX.get("http://fxtracker.test.cargill.com/User/Search")
    driverX.find_element_by_id("DSID").send_keys("t479091")
    driverX.find_element_by_id("searchButton").click()
    driverX.get("http://fxtracker.test.cargill.com/User/Edit/t479091")
    #driverX.get("http://fxtracker.test.cargill.com/User/Create/B383197")
    driverX.find_element_by_id("UserGroupID").click()
    time.sleep(1)
    ActionChains(driverX).key_down(Keys.DOWN).perform()
    time.sleep(1)
    ActionChains(driverX).key_down(Keys.DOWN).perform()
    time.sleep(1)
    ActionChains(driverX).key_down(Keys.DOWN).perform()
    time.sleep(1)
    ActionChains(driverX).key_down(Keys.DOWN).perform()
    driverX.find_element_by_id("BusinessUnitID").click()
    time.sleep(1)
    ActionChains(driverX).key_down(Keys.RETURN).perform()
    time.sleep(1)
    driverX.find_element_by_id("AddRole").click()
    driverX.get("http://fxtracker.test.cargill.com/user")
    time.sleep(2)
    driverX.close()




















########################## Writing in a file

with open(r'C:\Users\j417062\Desktop\Automation Py\Tickets.txt', 'w') as f:
    for item in df:
        f.write("%s\n" % item)
        
###################
        
dfObj = pd.DataFrame(output_rows) 

out_path = r'C:\Users\j417062\Desktop\Automation Py\Tickets.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
output_rows.to_excel(writer, sheet_name='Sheet1')
writer.save()

table_id = driver.find_element_by_id('T302087200').get_attribute('outerHTML') ## my tckets
df  = pd.read_html(table_id)

###########################


soup = BeautifulSoup(driver.find_element_by_id('T302087200').get_attribute('innerHTML'), 'html.parser')
table = soup.find("table")

df  = pd.read_html(soup)


table = pd.read_html(driver.find_element_by_id('WIN_6_303619800').get_attribute('innerHTML'))

output_rows = []
for table_row in table.findAll('tr'):
    columns = table_row.findAll('td')
    output_row = []
    for column in columns:
        output_row.append(column.text)
    output_rows.append(output_row)
    
with open(r'C:\Users\j417062\Desktop\Automation Py\Tickets.txt', 'w') as f:
    for item in output_rows:
        f.write("%s\n" % item)

with open(r'C:\Users\j417062\Desktop\Automation Py\Tickets.csv', 'wb') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(output_rows)
    
    type(output_rows)

#######################
tbody = table_id.find_element_by_tag_name("tr") # get all of the rows in the table
rows = tbody.find_elements_by_tag_name("tr")
name = 'Smith'

for row in rows:
    # Get the columns (all the column 2)        
    col = row.find_elements_by_tag_name("td") #note: index start from 0, 1 is col 2
    nobr = col.find_element_by_tag_name("nobr")
    span = nobr.find_elements_by_tag_name("span")
    print (span.text)
    #prints text from the element

list_rows = [[cell.text for cell in row.find_elements_by_tag_name('td')]
             for row in tbody.find_elements_by_tag_name('tr')]

for row in tbody.find_elements_by_tag_name('tr'):
    for cell in row.find_elements_by_tag_name('td'):
        span = cell.find_element_by_xpath("//span")
        print (span.text)

#######################


#pyautogui.moveTo(577, 1068)
#pyautogui.click()

#imX, imY = pyautogui.position()
imX = 101
imY = 484
pyautogui.moveTo(imX, imY)
pyautogui.click()
pyautogui.moveTo(imX, imY+2)
pyautogui.moveTo(266, 532)
pyautogui.click()
#select the group
############### wait 
#time.sleep(5) 
pyautogui.moveTo(474, 761)
pyautogui.click()
driver.find_element_by_id('arid_WIN_3_1000000217').send_keys("Apps-AMS-Trading-Gold-NA")
#ztime.sleep(2) 
pyautogui.moveTo(470, 778)
pyautogui.click()
driver.find_element_by_id('arid_WIN_3_1000000218').send_keys("JOHN SMITH")
pyautogui.moveTo(496, 863)
pyautogui.click()
pyautogui.moveTo(471, 914)
pyautogui.click()
#pyautogui.moveTo(412, 866)
#pyautogui.click()

driver.find_element_by_id('WIN_3_1002').click()



filterX, filterY = pyautogui.position()

print (filterX) 
print (filterY)


pyautogui.moveTo(filterX, filterY)
pyautogui.click()





