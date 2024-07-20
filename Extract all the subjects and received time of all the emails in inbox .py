import win32com.client
Outlook = win32com.client.Dispatch("Outlook.Application")
valores = Outlook.GetNamespace("MAPI").GetDefaultFolder(6)
for i in valores.Items:
    
    try:
            
        print(i.Subject," | ",i.ReceivedTime)
    

    except Exception as e:
        continue
