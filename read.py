import win32com.client

def ler_out():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)


    messages = inbox.Items
    message = messages.GetLast()

    print(message.SenderName)
    print(message.subject)
    msg = {message.body}


    print(msg)


