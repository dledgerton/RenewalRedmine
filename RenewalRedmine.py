
from redmine import Redmine
import settings
import xlsxwriter

redmine = Redmine(settings.SITE,
                  username=settings.USERNAME,
                  password=settings.PASSWORD,
                  requests={'verify':False}
                  )

# Open file for writing
    
f = open('c:/users/a62478/Desktop/data.txt','w',encoding='utf-8')
    
# Create the header row for the text file
    
header = 'ISSUE_ID|SUBJECT|DESCRIPTION|AUTHOR|CREATED_ON|ASSIGNED_TO|ESTIMATED_HOURS|PRIORITY|START_DATE|STATUS|TRACKER|UPDATED_ON'
    
f.write(header)
f.write('\n')

#Add all project information to the text file

RedmineProjects = ['baan','mes','pcf']

for project in RedmineProjects:

    selected_project = redmine.project.get(project)

    for issue in selected_project.issues:
    
    # if the object has the attribute print it
    
        if hasattr(issue, 'id'):
            line = str(issue.id)
        else:
            line = 'NA'
    
        if hasattr(issue, 'subject'):
            line = line + '|' + str(issue.subject)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'description'):
            line = line + '|' + str(issue.description).replace('\n', ' ').replace('\r', '')
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'author'):
            line = line + '|' + str(issue.author)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'created_on'):
            line = line + '|' + str(issue.created_on)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'assigned_to'):
            line = line + '|' + str(issue.assigned_to)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'estimated_hours'):
            line = line + '|' + str(issue.estimated_hours)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'priority'):
            line = line + '|' + str(issue.priority)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'start_date'):
            line = line + '|' + str(issue.start_date)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'status'):
            line = line + '|' + str(issue.status)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'tracker'):
            line = line + '|' + str(issue.tracker)
        else:
            line = line + '|' + 'NA'
    
        if hasattr(issue, 'updated_on'):
            line = line + '|' + str(issue.updated_on)
        else:
            line = line + '|' + 'NA'
    
        f.write(line + '\n')
    
f.close()

