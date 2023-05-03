# Production_automation_project

# About : This Project is automation of daily task of validating 400+ URLs(websites/portal)

# Flow :
The script takes the URLs input from the excel sheet.
It than iterates to all the URL and loads it in the chrome browser in a loop.
It than wait for the URLs to be completely loaded.
It than takes a full page screenshot of the page/website and cover 100% of the portal.
It than check if the current page it took screenshot of matches with the expected output.
This expected output is also the reuslt of the initial dry run of the code where expected output is stored with the script.
It than matches the current result with the expected one.
If the match is 100% than it adds the reult in the new excel sheet in the adjacent row of the URL as pass with green color
If there is any failure of even 0.1% and the result doesnt match than it updates the excel adjacent to the URL as failed URL.
It iterates to all the mentioned URL similarly and save the result in the excel sheet.
Once all the URLs are properly checked it than generates a report of the excel and sends auto mail with the reports excel over the mail and notifes the failing URLs.
Action can be taken based on thi reports.
The report is than send to the concerned detail for daily Report.

# Production Use:
It reduces the time by 95% and thereby reducing the cost as well.
It reduces the manual efforts by upto 98%.
It increases the efficency by upto 80% of the check being performed.
It improves the reporting and anaysis of the reports and auto sends the mail for reporting.
It is helpfull to reduce the cost, efforts, time thereby increasing the efficency.
