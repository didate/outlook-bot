import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

mildas = []
for message in messages:
    if message.Subject =='Workspace ONE UEM - Device Successfully Enrolled':
        date = message.ReceivedTime.strftime("%m/%d/%Y")
        if(date == '04/11/2022' or date == '04/12/2022' or date == '04/13/2022' or date == '04/14/2022'): 
            milda_index = message.body.index('MILDA')
            milda = message.body[milda_index:milda_index+9]
            mildas.append(str(message.ReceivedTime) +","+milda)


resultFyle = open("milda2.csv",'w')
for r in mildas:
    resultFyle.write(r + "\n")
resultFyle.close()