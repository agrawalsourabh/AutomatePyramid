
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
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


# Timesheet
driver.get("https://pyramidcore.pyramidci.com/TS/PCItimesheet.aspx")
title = driver.title
print(title)


# Date - txtDate
date_element = driver.find_element_by_id("txtDate")
print(str(date_element.get_attribute('value')))

# Master Project - ddlMstrProj
master_project_element = driver.find_element_by_id("ddlMstrProj")
select = Select(master_project_element)
print (select.first_selected_option.text)

# Project - ddlProject
project_element = driver.find_element_by_id("ddlProject")
select = Select(project_element)
print(select.first_selected_option.text)

# Features - ddlFeature
features_element = driver.find_element_by_id("ddlFeature")
select = Select(features_element)
print(select.first_selected_option.text)

# Ticket No - ddlTicketNo
ticket_no_element = driver.find_element_by_id("ddlTicketNo")
select = Select(ticket_no_element)
print(select.first_selected_option.text)

# Group - ddlGroup
group_element = driver.find_element_by_id("ddlGroup")
select = Select(group_element)
print(select.first_selected_option.text)

# Activity - ddlActivity
activity_element = driver.find_element_by_id("ddlActivity")
select = Select(activity_element)
print(select.first_selected_option.text)

# hours - ddlHour
hours_element = driver.find_element_by_id("ddlHour")
select = Select(hours_element)
print(select.first_selected_option.text)

# Add button
add_element = driver.find_element_by_id("btnAdd")
