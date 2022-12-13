import win32com.client as win32
import datetime
import pathlib

import sys

from custom_exceptions import FolderNotFoundException

# setting path
sys.path.append('..')



from typing import Optional, Callable, Type




def print_folder_names(outlook, account='ppzmis@exmail.nottingham.ac.uk'):
    """
    Quick utility method to show folder names. Construct
    Hierarchy of names in tuple as input to get_emails.
    """
    inbox = outlook.Folders('Inbox')
    for folder in inbox.Folders:
        #index starts from 1
        print(folder)
        for sub_folder in folder.Folders:
            print('\t' + str(sub_folder))
            for sub_sub_folder in sub_folder.Folders:
                print('\t\t' + str(sub_sub_folder))

def open_outlook(account='ppzmis@exmail.nottingham.ac.uk'):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders(account)
    return outlook

def get_emails(outlook, folder: tuple=tuple('Inbox'), filter : dict={}) -> list[Type[win32.CDispatch]]:
    """A function to scan emails on local Outlook client

    Args:
        folder (tuple): Specifies folder tree. Tuple of strings starting with top level folder descending to lowest subfolder. Call print_folder_names to see what is what
        filter (dict): a dictionary specifying filters to the emails returned. Defaults to {}

    All parameters in filter are optional. parameters which take str check if the str is contained in subject etc. Takes wildcards

    filter = {
                'start': Union[str dd/mm/yy, datetime.datetime],
                'stop': Union[str dd/mm/yy, datetime.datetime],
                'sender_email': str,
                'subject': str,
                'body': str,
                'has_attachments': bool
             }

    Returns:
        list[Type[win32com.client.CDispatch]]: A list of all emails which match criteria in filter_fn
    """
    
    mailbox = outlook
    try:
        for folder_name in folder:
            mailbox = mailbox.Folder(folder_name)
    except e:
        raise FolderNotFoundException(e)
    messages=mailbox.Items

    if 'start' in filter.keys():
        messages = messages.Restrict("[ReceivedTime] >= '" + filter['start_date'])
    if 'stop' in filter.keys():
        messages = messages.Restrict("[ReceivedTime] <= '" + filter['stop_date'])
    if 'sender_email' in filter.keys():
        filter = "@SQL=" & chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " ci_phrasematch " + filter['sender_email'] 
        messages = messages.Restrict(filter)
        
    #filtered_messages: list[Type[win32com.client.CDispatch]] = filter#(messages)
    return messages


def move_emails(outlook : Type[win32.Dispatch], messages : list, new_folder : tuple=('Inbox')):
    """Moves messages to new folder

    Args:
        outlook (Type[win32.Dispatch]): Mail Object
        messages (list): list of message items returned by get_emails
        new_folder (tuple, optional): tuple of strs corresponding to hierarcy of folders. Defaults to ('Inbox').

    Raises:
        FolderNotFoundException: _description_
    """

    to_mailbox = outlook
    try:
        for folder_name in new_folder:
            to_mailbox = to_mailbox.Folder(folder_name)
    except e:
        raise FolderNotFoundException(e)

    for message in messages:
        message.move(to_mailbox)


def send_email(msg: dict, attachments=None):
    """Sends an email using the local outlook

    Args:
        msg (dict): {
                        'to'        : email / emails of recipients,
                        'subject'   : subject of email,
                        'body'      : message in email,
                        'html_body' : message with html formatting,
                    }
        attachments (_type_, optional): _description_. Defaults to None.
    
        'to' and 'subject' required, 
        'body' or 'html_body' are optional
        'html_body' only applied if 'body' not present in msg.keys()
    """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = msg['to']
    mail.Subject = msg['subject']
    
    if 'body' in msg.keys():
        mail.Body = msg['body']
    elif 'html_body' in msg.keys():
        mail.HTMLBody = msg['html_body']

    if attachments is not None:
        for attachment in attachments:
            mail.Attachments.Add(attachment)

    mail.Send()


if __name__ == '__main__':
    outlook = open_outlook()
    print_folder_names(outlook)

    msg = {'to':'mike.i.smith@nottingham.ac.uk',
            'subject':'Test',
            'html_body':'<h1>Test</h1>'}

    send_email(msg)