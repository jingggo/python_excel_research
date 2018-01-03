import win32com.client as win32
import sqlite3
import xlrd,datetime,xlwt

def sendEmail():
    '''send email'''
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    receivers = ['yqian@7thonline.cn']
    mail.To = receivers[0]
    mail.Subject='test2'
    workbook = xlrd.open_workbook('simple.xls')
    mySheet = workbook.sheet_by_index(0)

    nrows = mySheet.nrows
    content = []

    for i in range(nrows):
        ss = mySheet.row_values(i)
        content.append(ss)
        print(content)
        Truecontent = str(content)

    mail.Body = 'test2'
    mail.Attachments.Add(r'D:\learngitt\python_excel_research\simple.xls')
    mail.Send()

def receiveEmail():
    '''start/run outlook'''
    outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
    '''get index instance'''
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    print 'total message', len(messages)
    # message = messages.GetLast()
    i = 0
    for message in messages:
        print message.Subject
        i += 1

    # infos=[]
    #
    # if hasattr(message,"Subject"):
    #     subject = message.Subject
    #     infos.append(subject)
    # if hasattr(message,"ReceivedTime"):
    #     received_time=str(message.ReceivedTime)
    #     # received_time=datetime.strptime(received_time,"%m/%d/%y %H:%M:%S")
    #     infos.append(received_time)
    # if hasattr(message,"HTMLBody"):
    #     html_body=message.HTMLBody
    #     infos.append(html_body)
    #     size=long(message.Size)
    # if hasattr(message,"SenderName"):
    #     sender=message.SenderName
    #     infos.append(sender)
    # if hasattr(message,"To"):
    #     receiver=message.To
    #     infos.append(receiver)
    # if hasattr(message,"Cc"):
    #     cc=message.Cc
    #     infos.append(cc)
    # if hasattr(message,"Body"):
    #     body=message.Body
    #     infos.append(body)
    #
    # print(infos)

receiveEmail()