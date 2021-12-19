import win32com.client as client


outlook = client.Dispatch("Outlook.Application")


namespace = outlook.GetNameSpace("MAPI")
inbox = namespace.GetDefaultFolder(6)


target_subject = "Тестовая тема 1"


mail_items = [item for item in inbox.Items if item.Class == 43]


for item in mail_items:
    if item.Subject != target_subject:
        message = item.Reply()
        message.Body = "Ошибка в теме письма"
        message.Save()
        message.Send()
    else:
        if item.Attachments.Count > 0:
            attachments = item.Attachments
            isFind = False
            i = 1
            fileList = []
            for file in attachments:
                if file.FileName[file.FileName.rfind("."):] == ".pdf":
                    isFind = True
                    fileList.append(file)
                    i += 1
            if not isFind:
                message = item.Reply()
                message.Body = "Отсутствует необходимый файл вложения"
                message.Save()
                message.Send()
            else:
                if i == 1:
                    fileList[0].SaveAsFile("C:\\Users\\yakhi\\Рабочий стол\\Тест\\{}".format(item.Subject + " - " + str(item.ReceivedTime).replace(":", " ") + ".pdf"))
                else:
                    fileList[0].SaveAsFile("C:\\Users\\yakhi\\Рабочий стол\\Тест\\{}".format(item.Subject + " - " + str(item.ReceivedTime).replace(":", " ") + " " + str(i) + ".pdf"))
        else:
            message = item.Reply()
            message.Body = "Отсутствует необходимый файл вложения"
            message.Save()
            message.Send()