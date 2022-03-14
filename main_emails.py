import pandas as pd
import os
import datetime
import win32com.client

csvFilePath = r".\kickstarter_projects.csv"
dfCsvData = pd.read_csv(csvFilePath)

# print(dfCsvData)

dfCsvData = dfCsvData.applymap(lambda x: x.encode('unicode_escape').
decode('utf-8') if isinstance(x, str) else x)

# print(dfCsvData)

dfCsvData['Launched'] = pd.to_datetime(dfCsvData['Launched'], format='%d-%m-%Y').dt.date
dfCsvData['Deadline'] = pd.to_datetime(dfCsvData['Deadline'], format='%d-%m-%Y').dt.date

# print(dfCsvData)
# print(dfCsvData['Launched'])

dfUniqueCountries = sorted(dfCsvData['Country'].unique())

# print(dfUniqueCountries)

tempFolder = "temp"
os.makedirs(tempFolder, exist_ok = True)
for each_country in dfUniqueCountries:
  tempStore = dfCsvData.loc[dfCsvData['Country'] == each_country]
  fileName = each_country + ".xlsx"
tempStore.to_excel(os.path.join(tempFolder, fileName),sheet_name=each_country, index = False, header=True)

fileSize = os.path.getsize(os.path.join(tempFolder, fileName))
print(fileSize)

if fileSize < 20971520:
  outlook = win32com.client.Dispatch('outlook.application')
  mail = outlook.CreateItem(0)
  mail.To = 'arpanghoshsharp@gmail.com'
  mail.Subject = 'Test Email with Attachment for country - ' + each_country
  mail.HTMLBody = 'Hello,<br><br>' + 'The email contains a test attachment for country - ' + each_country + '.' + '<br><br>' + 'Thanks,<br>' + 'Mohan'
  mail.Attachments.Add(os.path.join(tempFolder, fileName))
  #mail.CC = 'somebody.cc@newcompany.com'
  mail.Save()
  if os.path.exists(os.path.join(tempFolder, fileName)): os.remove(os.path.join(tempFolder, fileName))
else:
  os.remove(os.path.join(tempFolder, fileName))

os.rmdir(tempFolder)
print("Done")