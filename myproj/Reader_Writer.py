import pandas as pd
from datetime import date 

def log_in_excel(timesheet_date, master_project, project, features, ticketNo, group, activity, hours, isSaved, isSubmitted) :
    reports = pd.read_excel('Report.xlsx')
    row = reports.shape[0] +  1
    current_date = date.today()

    new_entry = {
        'Sno.'              : row,
        'Script Run Date'   : current_date,
        'Timesheet Date'    : timesheet_date,
        'Master Project'    : master_project,
        'Project'           : project,
        'Features'          : features,
        'TicketNo'          : ticketNo,
        'Group'             : group,
        'Activity'          : activity,
        'hours'             : hours,
        'isSaved'           : isSaved,
        'isSubmitted'       : isSubmitted          
    }

    reports.append(new_entry, ignore_index=True)
    reports.to_excel("Report.xlsx")

    return 1