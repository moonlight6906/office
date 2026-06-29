import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

items = inbox.Items
items.Sort("[ReceivedTime]", True)

for item in items:
    try:
        if item.MessageClass.startswith("IPM.Note") and item.UnRead:
            print("Subject:", item.Subject)
            print("From:", item.SenderName)
            print("Received:", item.ReceivedTime)
            print("Body:", item.Body)
            break
    except Exception:
        pass
