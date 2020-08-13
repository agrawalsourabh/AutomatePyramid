from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from BaseScript import setup

import time

from Reader_Writer  import log_in_excel

def fill_attendance():
    # Attendance
    driver = setup()
    driver.get("https://pyramidcore.pyramidci.com/Reports/PCI_HRReport.aspx")
    title = driver.title
    print(title)

    # update attendance
    update_attendance_bn = driver.find_element_by_id("btnDiscrepency")
    update_attendance_bn.click()

    time.sleep(2)

    table = driver.find_elements_by_xpath("//table[@id='dgIssueList']")
    row_count = len(driver.find_elements_by_xpath("//table[@id='dgIssueList']/tbody/tr"))
    column_count = len(driver.find_elements_by_xpath("//table[@id='dgIssueList']/tbody/tr/td"))

    print("Rows: " + str(row_count))
    print("Column: " + str(column_count))
    
    if(row_count == 0):
        raise Exception("Row number starts from 1")

    row_count = row_count + 1
    row = table[1].find_elements_by_xpath("//tr["+str(row_count)+"]/td")
    print(len(row))
    # rData = []
    # for webElement in row :
    #     rData.append(webElement.text)

    # print(rData)



if __name__ == "__main__":
    fill_attendance()
