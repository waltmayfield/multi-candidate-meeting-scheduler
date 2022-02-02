# import win32com.client
# import csv

# outlook = win32com.client.Dispatch('Outlook.Application')
# namespace = outlook.GetNamespace("MAPI")

# recipient = namespace.CreateRecipient('waltmayf@amazon.com')



# # print(list(dir(recipient)))

# air_port_code = recipient.AddressEntry.GetExchangeUser().OfficeLocation[:3]

# def air_port_code_to_tz(air_port_code,filename = 'tzmap.txt'):
# 	with open(filename, 'r') as file:
# 		datareader = csv.reader(file, delimiter='\t')
# 		for code, timezone in datareader:
# 		        if code == air_port_code: return timezone


# print(recipient.AddressEntry.GetExchangeUser().OfficeLocation[:3])

# print(air_port_code_to_tz(air_port_code))
print('\u2022')