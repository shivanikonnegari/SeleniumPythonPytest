import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import XLUtilities

driver = webdriver.Chrome()
driver.get("https://www.letskodeit.com/login")
driver.maximize_window()

path = "C://selenium performance/LoginTest (1).xlsx"
time.sleep(5)

rows = XLUtilities.getRowCount(path, "Sheet1")

for r in range(2,rows+1):
    username = XLUtilities.readData(path, "Sheet1",r, 1)
    password = XLUtilities.readData(path,"Sheet1",r, 2)

    driver.find_element(By.ID,"email").send_keys(username)
    time.sleep(5)
    driver.find_element(By.NAME, "password").send_keys(password)
    time.sleep(5)
    driver.find_element(By.XPATH,"//button[@id='login']").click()
    time.sleep(4)

    driver.find_element(By.XPATH,"//img[@class='zl-navbar-rhs-img ']").click()
    driver.find_element(By.XPATH,"//a[normalize-space()='Logout']").click()

    if driver.title == "Home Page":
        print("Test is Passed")
        XLUtilities.writeData(path,  "Sheet1", r,   3, "Passed")

    else:
        print("Test is Failed")
        XLUtilities.writeData(path, "Sheet1", r, 3, "Failed")

    driver.find_element(By.XPATH, "//a[normalize-space()='Sign In']").click()
    time.sleep(4)