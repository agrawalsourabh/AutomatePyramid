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
    
    # //*[@id="dgIssueList"]/tbody/tr[2]/td[2]

    for r in range(row_count):

        if r > 0:
            date_xpath = "//*[@id='dgIssueList']/tbody/tr[" + str(r+1) + "]/td[2]"
            row = driver.find_element_by_xpath(date_xpath)
            date_text = row.text
            print(date_text)

            if r < 9:
                in_date_id = "dgIssueList_ctl0" + str(r+1) + "_txtDateIn"
                out_date_id = "dgIssueList_ctl0" + str(r+1) + "_txtDateOut"

            else:
                in_date_id = "dgIssueList_ctl" + str(r+1) + "_txtDateIn"
                out_date_id = "dgIssueList_ctl" + str(r+1) + "_txtDateOut"


            in_date = driver.find_element_by_id(in_date_id)
            in_date.send_keys(date_text)

            out_date = driver.find_element_by_id(out_date_id)
            out_date.send_keys(date_text)

        
    
    

if __name__ == "__main__":
    fill_attendance()
