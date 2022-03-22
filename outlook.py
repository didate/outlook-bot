import win32com.client
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

for message in messages:
    if message.Subject =='Workspace ONE UEM Device Activation':
        token_index = message.body.index('Token: ')
        token = message.body[token_index+7:token_index+13]
        date = message.ReceivedTime.strftime("%d/%m/%Y")
        if(date == '22/03/2022'):
            print(token)
