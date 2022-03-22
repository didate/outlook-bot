import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

tokens = []
for message in messages:
    if message.Subject =='Workspace ONE UEM Device Activation':
        date = message.ReceivedTime.strftime("%d/%m/%Y")
        if(date == '22/03/2022'): 
            token_index = message.body.index('Token: ')
            token = message.body[token_index+7:token_index+13]
            tokens.append(token)


resultFyle = open("output.csv",'w')
for r in tokens:
    resultFyle.write(r + "\n")
resultFyle.close()