import time
from fileinput import filename

from selenium import webdriver
from selenium.webdriver.common.by import By

import XLUtilities

driver=webdriver.Chrome()
driver.get("https://only-testing-blog.blogspot.com/2014/05/form.html")
driver.maximize_window()
time.sleep(5)

path = "C://selenium performance/DataDrivenTesting.xlsx"

rows=XLUtilities.getRowCount(path,"Sheet1")

for r in range(2,rows+1):
    firstname=XLUtilities.readData(path,'Sheet1',r,1)
    lastname = XLUtilities.readData(path, "Sheet1",r,2)
    email = XLUtilities.readData(path, 'Sheet1', r, 3)
    mobileno= XLUtilities.readData(path, "Sheet1", r, 4)
    companyname= XLUtilities.readData(path, 'Sheet1', r, 5)


    driver.find_element(By.NAME,'FirstName').send_keys(firstname)
    driver.find_element(By.NAME, 'LastName').send_keys(lastname)
    driver.find_element(By.NAME, 'EmailID').send_keys(email)
    driver.find_element(By.NAME, 'MobNo').send_keys(mobileno)
    driver.find_element(By.NAME, 'Company').send_keys(companyname)
    driver.find_element(By.XPATH, '//*[@id="post-body-8228718889842861683"]/div[1]/form/input[6]').click()
    time.sleep(5)

    driver.switch_to.alert.accept()
    time.sleep(5)

    if driver.title=="Only Testing: Form":
        print("test is passed")
        XLUtilities.writeData(path,"Sheet1",r,6,"Passed")
    else:
        print("test is failed")
        XLUtilities.writeData(path, "Sheet1", r, 6, "Failed")