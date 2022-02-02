#https://stackoverflow.com/questions/20222941/retrieving-free-busy-status-from-microsoft-outlook-using-python


import win32com.client
import datetime
import os
import subprocess
import time
# import xlrd

# sheetLocation = os.path.abspath(os.getcwd()+"\seat.xlsx")

# #Open Workbook 
# wb = xlrd.open_workbook(sheetLocation) 
# sheet = wb.sheet_by_index(0) 
# len = sheet.nrows


# time.sleep(2)
# os.system("taskkill /f /im outlook.exe")    #Kill outlook application if open
# time.sleep(1)
openOutlook = win32com.client.Dispatch("Outlook.Application")
mail = openOutlook.CreateItem(1)   #Open new mail

#Adds 2 days for the current date
#The 'days' field can be set to  anything i.e, future date or past date all wrt current date
todayDate = datetime.date.today()+ datetime.timedelta(days=2)

#Convert the int returned to string to append time
todayDateStr=str(todayDate)

#Start Time 
startTime =" 12:00"

print((todayDateStr+startTime))
mail.Start = (todayDateStr+startTime)  #Append start time to the 
mail.Subject = 'Accidental_Send_Please_Delete'
mail.Duration = 15                     #Duration in minutes
mail.MeetingStatus = 1                  #Availability

mail.Recipients.Add("waltmayf@amazon.com")


# #Iterate over all the participants and add them to the "To"
# for i in range(len):    
#     mail.Recipients.Add(sheet.cell_value(i,0)) 

# mail.Save()
mail.Display()
# mail.Send()

# subprocess.call(["C:/Program Files (x86)/Microsoft Office/root/Office16/OUTLOOK.EXE"])

# os._exit(0)