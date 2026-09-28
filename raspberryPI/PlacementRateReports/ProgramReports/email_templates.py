from datetime import date

def month_end_cd(contact_name, year):
    return (
f"Good Morning {contact_name},\n\n"
"Today is the last day of the month. The report has been updated with our current, month-end placement statistics.\n\n"
f"The first excel sheet tab contains the placement totals for the entire Class of {year}. "
"Each sheet after that contains data for your program or programs.\n"
"The sheets include a 'Most Recent Friday' table, which shows the data as of the day and time you received this email.\n"
"There is also a 'Weekly History' table, which provides a picture of how the placement numbers have been changing over the past couple of months.\n"
"Each table shows how many students are in each placement category. The most important categories are bolded. At the bottom of each table, you can see the placement percentage\n"
"If you have any questions, please contact the BCC Data Team. If there are any discrepancies in the data, please let us know.\n\n"
"Sincerely,\n"
"BCC Data Team"      
)

def week_end_cd(contact_name, year):
    return (
f"Good Morning {contact_name},\n\n"
"Here are the updated placement and internship reports for your programs from the past week.\n\n"
f"The first excel sheet tab contains the placement totals for the entire Class of {year}. "
"Each sheet after that contains data for your program or programs.\n"
"The sheets include a 'Most Recent Friday' table, which shows the data as of the day and time you received this email.\n"
"There is also a 'Weekly History' table, which provides a picture of how the placement numbers have been changing over the past couple of months.\n"
"Each table shows how many students are in each placement category. The most important categories are bolded. At the bottom of each table, you can see the placement percentage\n"
"The placement details sheet gives details about the specific student's who placed in the last week, Full-Time and Internship\n\n"
"If you have any questions, please contact the BCC Data Team. If there are any discrepancies in the data, please let us know.\n\n"
"Sincerely,\n"
"BCC Data Team"
)

def healthcare(contact_name):
    return (
f"Good Morning {contact_name},\n\n"
"Here are the updated placement and internship reports for your programs from the past week.\n\n"
"The first excel sheet tab contains all healthcare related Full Time placement details so far this year.\n"
"The second excel sheet tab contains all healthcare related Internship placement details so far this year.\n"
"The third excel sheet tab contains only last weeks placement details for both Full Time and Internship healthcare related placements.\n\n"
"If you have any questions, please contact the BCC Data Team. If there are any discrepancies in the data, please let us know.\n\n"
"Sincerely,\n"
"BCC Data Team"
)

def healthcare_monthend(contact_name):
    return (
f"Good Morning {contact_name},\n\n"
"Today is the last day of the month. Here are the updated placement and internship reports for your programs.\n\n"
"The first excel sheet tab contains all healthcare related Full Time placement details so far this year.\n"
"The second excel sheet tab contains all healthcare related Internship placement details so far this year.\n"
"The third excel sheet tab contains only last weeks placement details for both Full Time and Internship healthcare related placements.\n\n"
"If you have any questions, please contact the BCC Data Team. If there are any discrepancies in the data, please let us know.\n\n"
"Sincerely,\n"
"BCC Data Team"
)

def subject_and_body(contact_name, label, flag, today, year, has_visualization=False):
    monthly = flag in (1, 2)
    kind = 'Month End' if monthly else 'Weekly'
    subject = f'Class of {year}: {label} {kind} Placement Report {today}'
    if label == 'Healthcare':
        body = healthcare_monthend(contact_name) if monthly else healthcare(contact_name)
    else:
        body = month_end_cd(contact_name, year) if monthly else week_end_cd(contact_name, year)
    if has_visualization:
        body += f'\n\nThe attached PDF contains full-time placement dashboards as of {today}, generated from the attached report. Current-month values are snapshots as of this date. Prior-year comparisons retain their recorded observations.'
    return subject, body
