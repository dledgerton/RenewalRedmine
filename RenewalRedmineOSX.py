
from redmine import Redmine
import settings
import xlsxwriter

redmine = Redmine(settings.SITE,
                  username=settings.USERNAME,
                  password=settings.PASSWORD,
                  requests={'verify':False}
                  )

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('/Users/davidedgertonjr/desktop/RedmineData.xlsx')
worksheet = workbook.add_worksheet()

# Add an Excel date format.
date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

#Create header in the worksheet

categories = ['ISSUE_ID','SUBJECT','DESCRIPTION','AUTHOR','CREATED_ON','ASSIGNED_TO','ESTIMATED_HOURS','PRIORITY','START_DATE','STATUS','TRACKER','UPDATED_ON']
row = 0
col = 0

for item in categories:
    worksheet.write_string  (row, col,item )
    col += 1

#reset columns and rows to next available cell
col=0
row=1

#Add all project information to the text file

RedmineProjects = ['baan','mes','pcf']

for project in RedmineProjects:

    selected_project = redmine.project.get(project)

    for issue in selected_project.issues:
    
    # if the object has the attribute print it
    
        if hasattr(issue, 'id'):
            worksheet.write_number(row,col,issue.id)
        else:
            worksheet.write_string(row,col,'NA')
    
        if hasattr(issue, 'subject'):
            worksheet.write_string(row,col+1,issue.subject)
        else:
            worksheet.write_string(row,col+1,'NA')
    
        if hasattr(issue, 'description'):
            desc = str(issue.description).replace('\n', ' ').replace('\r', '')
            worksheet.write_string(row,col+2,desc)
        else:
            worksheet.write_string(row,col+2,'NA')
    
        if hasattr(issue, 'author'):
            worksheet.write_string(row,col+3,str(issue.author))
        else:
            worksheet.write_string(row,col+3,'NA')
    
        if hasattr(issue, 'created_on'):
            worksheet.write_datetime(row,col+4,issue.created_on,date_format)
        else:
            worksheet.write_string(row,col+4,'NA')
    
        if hasattr(issue, 'assigned_to'):
            worksheet.write_string(row,col+5,str(issue.assigned_to))
        else:
            worksheet.write_string(row,col+5,'NA')
    
        if hasattr(issue, 'estimated_hours'):
            worksheet.write_number(row,col+6,issue.estimated_hours)
        else:
            worksheet.write_string(row,col+6,'NA')
    
        if hasattr(issue, 'priority'):
            worksheet.write_string(row,col+7,str(issue.priority))
        else:
            worksheet.write_string(row,col+7,'NA')
    
        if hasattr(issue, 'start_date'):
            worksheet.write_datetime(row,col+8,issue.start_date,date_format)
        else:
            worksheet.write_string(row,col+8,'NA')
    
        if hasattr(issue, 'status'):
            worksheet.write_string(row,col+9,str(issue.status))
        else:
            worksheet.write_string(row,col+9,'NA')
    
        if hasattr(issue, 'tracker'):
            worksheet.write_string(row,col+10,str(issue.tracker))
        else:
            worksheet.write_string(row,col+10,'NA')
    
        if hasattr(issue, 'updated_on'):
            worksheet.write_datetime(row,col+11,issue.updated_on,date_format)
        else:
            worksheet.write_string(row,col+11,'NA')
        row +=1

workbook.close()