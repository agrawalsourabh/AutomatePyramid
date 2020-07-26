
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

exe_chrome_path = 'D:\personal\chromedriver_win32\chromedriver.exe'
driver = webdriver.Chrome(exe_chrome_path)

driver.get("https://pyramidcore.pyramidci.com/Security/PCILoginNew.aspx")
driver.maximize_window()
title = driver.title
print(title)

username_field = driver.find_element_by_id("pydLogin_txtUserid")
username_field.clear()
username_field.send_keys("sourabh.agrawal")

password_field = driver.find_element_by_id("pydLogin_txtUserPwd")
password_field.clear()
password_field.send_keys("Dec@2019")

login_btn = driver.find_element_by_id("pydLogin_btnLogin")
login_btn.click()

driver.get("https://pyramidcore.pyramidci.com/Home/PCIhome.aspx")

title = driver.title
print(title)