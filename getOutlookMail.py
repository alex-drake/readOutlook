# -*- coding: utf-8 -*-
"""
Created on Fri Dec 22 11:54:17 2017

@author: alexanderdrake
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Dec 21 08:23:35 2017

@author: alexanderdrake
"""
import site
sitedir = 'T:/IAT and Taskings/IT/Systems/Python/packages/'
site.addsitedir(sitedir)

import csv
import datetime
import win32com.client
import os
import PyPDF2

# connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# go to the TPOMA inbox
tpoma = outlook.Folders("INTEL").Folders("Inbox").Folders("INTELLIGENCE").Folders("FOR TPOMA")
msg = tpoma.Items
msgs = msg.GetFirst()

msgList = []
msgCnt = 0

while msgs:
    while msgCnt < msg.Count:
        # msg counter
        msgCnt = msgCnt + 1
        
        print(msgs.Subject) # user feedback
        
        # get sent date
        msgDate = msgs.SentOn
        msgDate = msgDate.strftime("%Y-%m-%d")    
        
        # check for attachments and save them if found
        atts = msgs.Attachments
        attsList = ''
        for i in range(atts.Count):
            attachment = atts.Item(i + 1)
            extension = os.path.splitext(str(attachment))[1][1:]
            filename = os.getcwd() + '\\attachments\\' + str(attachment)
            attsList = attsList + str(attachment) + ';'
            
            if extension in ['jpg','png','jpeg']:
                next 
            elif extension in ['pdf']:
                attachment.SaveAsFile(filename)
                pdfFileObj = open(filename, 'rb')
                pdfRead = PyPDF2.PdfFileReader(pdfFileObj)
                content = pdfRead.getFields()
                
                bodyText = ''
                for key in content.keys():
                    if key in ['From','To','Date & Time']:
                        next
                    else:
                        if '/V' in content[key]:
                            bodyText = bodyText + key + ': ' + str(content[key]['/V']) + ';'
                        else:
                            next
                # log each ASIF                                
                b2 = dict(
                        uniqueID = msgCnt,
                        #sentDate = content['Date & Time']['/V'],
                        sentDate = msgDate,
                        sentFrom = content['From']['/V'],
                        subject = 'ASIF',
                        body = bodyText,
                        att = 'NA'
                        )        
                msgList.append(b2)
                pdfFileObj.close()
            else:
                attachment.SaveAsFile(filename) # save attachments  
        
        # log each email
        b = dict(
        uniqueID = msgCnt,
        sentDate = msgDate,
        sentFrom = msgs.SenderName,
        subject = msgs.Subject,
        body = msgs.body,
        att = attsList
        )        
        msgList.append(b)
        
        msgs = msg.GetNext()
    else:
        break

# create filename (with path)  
exptName = os.getcwd() + '\\intel\\' + datetime.datetime.now().strftime("%Y%m%d") + '_intel.csv'
# export the intel as csv
with open(exptName,"w") as f:
    cw = csv.DictWriter(f,
                        fieldnames = msgList[0].keys(),
                        delimiter = ',',
                        lineterminator='\n')
    cw.writeheader()
    cw.writerows(msgList)
