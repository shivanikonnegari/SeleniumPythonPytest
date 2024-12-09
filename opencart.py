import time
from fileinput import filename
from selenium import webdriver
from selenium.webdriver.common.by import By


import XLUtilities
#step1:open the url on chrome browser
driver=webdriver.Chrome()
driver.get("https://demo.opencart.com.gr/")
driver.maximize_window()
time.sleep(5)

#step2:Verify if the title "Your Store" application is correct
actual_title = driver.title
expect_title = "Your Store"

if actual_title == expect_title:
    print("Title Verified")
else:
    print("Title Not Verified")

#step3:Click on "My Account" menu option
Myaccount=driver.find_element(By.XPATH,"/html/body/nav/div/div[2]/ul/li[2]/a/span[1]").click()
time.sleep(2)
#step4:select register option
Register=driver.find_element(By.XPATH,"/html/body/nav/div/div[2]/ul/li[2]/ul/li[1]/a").click()
#step5:Verify the text present on the web page as "Register Account"
actual_title = driver.title
expect_title = "Register Account"

if actual_title == expect_title:
    print("Webpage Title Verified")
else:
    print("Webpage Title Not Verified")

#step6:Enter all the details in the First name,Last name,E-mail,Telephone,Password,Confirm Password from the excel sheet(UserDetails.xls)

path="C://Users/Sigma/Book1 (1).xlsx"
rows=XLUtilities.getRowCount(path,"Sheet1")

for r in range(2,rows+1):
    FirstName=XLUtilities.readData(path,'Sheet1',r,1)
    LastName = XLUtilities.readData(path, "Sheet1",r,2)
    Email = XLUtilities.readData(path, 'Sheet1', r, 3)
    Telephone= XLUtilities.readData(path, "Sheet1", r, 4)
    Password= XLUtilities.readData(path, 'Sheet1', r, 5)
    ConfirmPassword = XLUtilities.readData(path, 'Sheet1', r, 6)

    driver.find_element(By.NAME,'firstname').send_keys(FirstName)
    driver.find_element(By.NAME, 'lastname').send_keys(LastName)
    driver.find_element(By.NAME, 'email').send_keys(Email)
    driver.find_element(By.NAME, 'telephone').send_keys(Telephone)
    driver.find_element(By.NAME, 'password').send_keys(Password)
    driver.find_element(By.NAME, 'confirm').send_keys(ConfirmPassword)


#step7:Select "I have read and agree to the Privacy Policy" check box
    newsletterRadiobtn = driver.find_element(By.NAME, "newsletter")
    newsletterRadiobtn.click()
    AgreeCheckbox = driver.find_element(By.NAME, "agree")
    AgreeCheckbox.click()
    time.sleep(2)

#step8:Click on "Continue" button
    driver.find_element(By.XPATH, "//input[@value='Continue']").click()
    time.sleep(5)

#step9:Verify the acknowledgement message "Your Account Has Been Created"

    if driver.title=="Your Account Has Been Created!":
        print("Test is passed")
        XLUtilities.writeData(path,"Sheet1",r,7,"Passed")
    else:
        print("Test is failed")
        XLUtilities.writeData(path, "Sheet1", r, 7, "Failed")
#closing browser
        driver.close()