from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from BaseScript import setup

import time

from Reader_Writer  import log_in_excel

def fill_timesheet():
    # Timesheet
    driver = setup()
    driver.get("https://pyramidcore.pyramidci.com/TS/PCItimesheet.aspx")
    title = driver.title
    print(title)



    # Master Project - ddlMstrProj
    master_project_element = driver.find_element_by_id("ddlMstrProj")
    select = Select(master_project_element)
    master_project_value = select.first_selected_option.text
    print (master_project_value)

    if master_project_value.lower() == 'weekend':
        submit_timesheet_for_weekend(driver, master_project_value.lower())

    else:
        for_weekday(driver, master_project_value.lower())


def for_weekday(driver, master_project):

    # timesheet_date, master_project, project, features, ticketNo, group, activity, hours, isSaved, isSubmitted
    
    # Date - txtDate
    date_element = driver.find_element_by_id("txtDate")
    timesheet_date = str(date_element.get_attribute('value'))
    print(timesheet_date)

    # Master Project - ddlMstrProj
    master_project_element = driver.find_element_by_id("ddlMstrProj")
    select = Select(master_project_element)
    select.select_by_visible_text('EMC')
    master_project = select.first_selected_option.text
    print (master_project)

    # Project - ddlProject
    time.sleep(3)
    project_element = driver.find_element_by_id("ddlProject")
    select = Select(project_element)
    select.select_by_visible_text('EMC-QA')
    project = 'EMC-QA'

    # Features - ddlFeature
    time.sleep(3)
    features_element = driver.find_element_by_id("ddlFeature")
    select = Select(features_element)
    select.select_by_visible_text('BOP')
    features = 'BOP'

    # Ticket No - ddlTicketNo
    time.sleep(3)
    ticket_no_element = driver.find_element_by_id("ddlTicketNo")
    select = Select(ticket_no_element)
    select.select_by_visible_text('Testing BOP')
    ticketNo = 'Testing BOP'

    # Group - ddlGroup
    time.sleep(3)
    group_element = driver.find_element_by_id("ddlGroup")
    select = Select(group_element)
    select.select_by_visible_text('Quality')
    group = 'Quality'

    
    time.sleep(3)
    # Activity - ddlActivity
    activity_element = driver.find_element_by_id("ddlActivity")
    select = Select(activity_element)
    select.select_by_visible_text('QA & Testing')
    print(select.first_selected_option.text)
    activity = select.first_selected_option.text

    hours_element = driver.find_element_by_id("ddlHour")
    select = Select(hours_element)
    select.select_by_visible_text('8')
    print(select.first_selected_option.text)
    hours = select.first_selected_option.text

    save_and_submit(driver, timesheet_date, master_project, project, features, ticketNo, group, activity, hours)
    


def submit_timesheet_for_weekend(driver, master_project):
    
    # timesheet_date, master_project, project, features, ticketNo, group, activity, hours, isSaved, isSubmitted

    # Date - txtDate
    date_element = driver.find_element_by_id("txtDate")
    timesheet_date = str(date_element.get_attribute('value'))
    print(timesheet_date)

    # Project - ddlProject
    project_element = driver.find_element_by_id("ddlProject")
    select = Select(project_element)
    project = select.first_selected_option.text
    print(project)

    # Features - ddlFeature
    features_element = driver.find_element_by_id("ddlFeature")
    select = Select(features_element)
    features = select.first_selected_option.text
    print(features)

    # Ticket No - ddlTicketNo
    ticket_no_element = driver.find_element_by_id("ddlTicketNo")
    select = Select(ticket_no_element)
    ticketNo = select.first_selected_option.text
    print(ticketNo)

    # Group - ddlGroup
    group_element = driver.find_element_by_id("ddlGroup")
    select = Select(group_element)
    group = select.first_selected_option.text
    print(group)

    # Activity - ddlActivity
    activity_element = driver.find_element_by_id("ddlActivity")
    select = Select(activity_element)
    activity = select.first_selected_option.text
    print(activity)

    # hours - ddlHour
    hours_element = driver.find_element_by_id("ddlHour")
    select = Select(hours_element)
    hours = select.first_selected_option.text
    print(hours)

    save_and_submit(driver, timesheet_date, master_project, project, features, ticketNo, group, activity, hours)


def save_and_submit(driver, timesheet_date, master_project, project, features, ticketNo, group, activity, hours):
    # Add button
    add_element = driver.find_element_by_id("btnAdd")
    add_element.click()

    # check successfully message - lblmsg
    wait = WebDriverWait(driver, 30)
    element = wait.until(EC.element_to_be_clickable((By.ID, 'lblmsg')))
    message_element = driver.find_element_by_id("lblmsg")
    message = message_element.text

    isSaved = 'N'

    print(message)
    if 'successfully'  in message.lower():
        isSaved = 'Y'

    else:
        return

    # Save button - btnSave
    wait = WebDriverWait(driver, 30)
    element = wait.until(EC.element_to_be_clickable((By.ID, 'btnSave')))
    save_btn = driver.find_element_by_id("btnSave")
    save_btn.click()



    # Submit Button - btnSubmit
    submit_btn = driver.find_element_by_id("btnSubmit")
    submit_btn.click()

    # Handling alert
    alert_obj = driver.switch_to.alert
    alert_obj.accept()

    # Read message after submiting the timesheet.
    wait = WebDriverWait(driver, 30)
    element = wait.until(EC.element_to_be_clickable((By.ID, 'lblmsg')))
    message_element = driver.find_element_by_id("lblmsg")
    message = message_element.text
    print(message_element.text)

    isSubmitted = 'N'

    if 'successfully'  in message.lower():
        isSubmitted = 'Y'

    islogged = log_in_excel(timesheet_date, master_project, project, features, ticketNo, group, activity, hours, isSaved, isSubmitted)
    print(islogged)


if __name__ == "__main__":
    fill_timesheet()