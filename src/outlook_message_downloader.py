"""
Created on Wed May 19 21:00:09 2021

@author: rledezmah071

Useful links: 
https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook/35801030#35801030
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?redirectedfrom=MSDN&view=outlook-pia#properties_
"""

import win32com.client
import pandas as pd
import datetime

def get_outlook_messages(root_folder, folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    root = mapi.Folders[root_folder]
    folder_to_use = root.Folders[folder_name]
    messages = folder_to_use.Items
    return messages
    

def get_message_attributes(messages, min_date, max_date):
    messages.Sort("[ReceivedTime]", True)
    
    min_date_sfilter = "[ReceivedTime] >= '"\
                 + datetime.datetime.strptime(min_date , '%Y-%m-%d')\
                     .strftime('%m/%d/%Y %H:%M %p') + "'"
    max_date_sfilter = "[ReceivedTime] < '"\
                 + datetime.datetime.strptime(max_date , '%Y-%m-%d')\
                     .strftime('%m/%d/%Y %H:%M %p') + "'"
                     
    min_filtered_messages = messages.Restrict(min_date_sfilter)
    filtered_messages = min_filtered_messages.Restrict(max_date_sfilter)
                
    sender_list = []
    time_list = []
    subject_list = []
    body_list = []

    for message in filtered_messages:
        if message.Class == 43:
            try:
                if message.SenderEmailType == 'EX':
                    sender_list.append(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                else:
                    sender_list.append(message.SenderEmailAddress)
            except:
                sender_list.append('N/A')            
            try:
                time_list.append(str(message.ReceivedTime.date()))
            except:
                time_list.append('N/A')
            try:
                subject_list.append(message.Subject)
            except:
                subject_list.append('N/A')
            try:
                body_list.append(message.Body)
            except:
                body_list.append('N/A')
    
    dataframe = pd.DataFrame(list(zip(sender_list, time_list, subject_list, body_list)),
               columns =['sender', 'date', 'subject', 'body'])
    
    return dataframe
    
