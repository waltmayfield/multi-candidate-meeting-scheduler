#https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder
#https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/create-an-appointment-as-a-meeting-on-the-calendar
#https://pythoninoffice.com/get-outlook-calendar-meeting-data-using-python/

import win32com.client

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

# for item in outlook.Items: print(item)

# print(outlook)

