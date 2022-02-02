#https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace.createrecipient
#https://github.com/iakov-aws/slots/blob/55bec95b9d1be7ec112af18b1311d35a2a277fbc/slots.py
#Lol this amazonian has the perfect solution.


import win32com.client
import datetime

obj_outlook = win32com.client.Dispatch('Outlook.Application')
obj_Namespace = obj_outlook.GetNamespace("MAPI")
obj_Recipient = obj_Namespace.CreateRecipient("waltmayf@amazon.com")
# my_date = datetime.date(2021,2,20)
str_Free_Busy_Data = obj_Recipient.FreeBusy(my_date, 11)
# print(str_Free_Busy_Data)
print(type(obj_Recipient))
