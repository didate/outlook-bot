import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

tokens = []
count = 0
row = ""
for message in messages:
    if message.Subject =='Workspace ONE UEM Device Activation':
        date = message.ReceivedTime.strftime("%d/%m/%Y")
        if(date == '07/05/2022'): 
            token_index = message.body.index('Token: ')
            token = message.body[token_index+7:token_index+13]
            row = row + "," + token
            count=count+1
            if(count == 22):
                print(row)
                tokens.append(row)
                count =0
                row = ""

if row != "":
    tokens.append(row)

resultFyle = open("output20220507_.csv",'w')
for r in tokens:
    resultFyle.write(r + "\n")
resultFyle.close()