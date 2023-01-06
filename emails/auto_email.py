from pydates.pydates import parse_date, format_datetime_to_str

import win32com.client as win32
import datetime
import pathlib
import time
from typing import Optional, Callable, Type, Dict
import uuid

# setting path
from emails.custom_exceptions import FolderNotFoundException, EmailAttachmentException




def open_outlook(account='ppzmis@exmail.nottingham.ac.uk'):
    """Create new instance of Outlook

    Args:
        account (str, optional): account email. Defaults to 'ppzmis@exmail.nottingham.ac.uk'.

    Returns:
        _type_: Outlook Object's folder associated with email
    """
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders(account)
    return outlook

def find_folder(outlook, folder: tuple=('Inbox',)) -> Type[win32.CDispatch]:
    """Navigates outlook email folder structure and returns object containing all message objects

    Args:
        outlook (_type_): outlook instance
        folder (tuple, optional): folder tuple e.g ('Inbox', 'Admin', 'Church'). Defaults to ('Inbox',).

    Raises:
        FolderNotFoundException: simply raises and prints description of error

    Returns:
        Type[win32.CDispatch]: mailfolder object
    """
    mailbox = outlook
    try:
        for folder_name in folder:
            mailbox = mailbox.Folders(folder_name)
    except Exception as e:
        raise FolderNotFoundException(e)

    return mailbox

def get_emails(outlook, folder: tuple=('Inbox',), filter : dict={}) -> list[Type[win32.CDispatch]]:
    """A function to scan emails on local Outlook client

    Args:
        folder (tuple): Specifies folder tree. Tuple of strings starting with top level folder descending to lowest subfolder. Call print_folder_names to see what is what
        filter (dict): a dictionary specifying filters to the emails returned. Defaults to {}

    All parameters in filter are optional. parameters which take str check if the str is contained in subject etc.
    has_attachments will only return emails with attachments if True but all emails if False
    It is not recommended to use from_email and from. Likely to produce weird results.

    filter = {
                'start': Union[str dd/mm/yy, datetime.datetime],
                'stop': Union[str dd/mm/yy, datetime.datetime],
                'from_email': str,
                'from_name : str,
                'cc_email': str.
                'subject': str,
                'body': str,
                'html_body':str,
                'has_attachments': bool
             }

    Returns:
        list[Type[win32com.client.CDispatch]]: A list of all emails which match criteria in filter_fn
    """

    messages = find_folder(outlook, folder=folder).Items

    #Apply filters to messages
    #https://docs.oracle.com/cd/E13218_01/wlp/compozearchive/javadoc/portlets20/com/compoze/exchange/webdav/HttpMailProperty.html
    if 'start' in filter.keys():
        #Note format string is American Y-d-m for the benefit of Outlook
        messages = messages.Restrict("[ReceivedTime] >= '" + format_datetime_to_str(parse_date(filter['start']),format="%Y-%d-%m %H:%M %p") + "'")
    if 'stop' in filter.keys():
        messages = messages.Restrict("[ReceivedTime] <= '" + format_datetime_to_str(parse_date(filter['stop']),format="%Y-%d-%m %H:%M %p") + "'")
    if 'from_email' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:fromemail Like '%" + filter['from_email'] + "%'" )
    if 'from_name' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:fromname Like '%" + filter['from_name'] + "%'" )
    if 'cc_email' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:cc Like '%" + filter['sender_email'] + "%'" )
    if 'subject' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:subject Like '%" + filter['subject'] + "%'" )
    if 'body' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:textdescription Like '%" + filter['body'] + "%'" )
    if 'html_body' in filter.keys():
        messages = messages.Restrict("@SQL=urn:schemas:httpmail:htmldescription Like '%" + filter['html_body'] + "%'" )
    if 'has_attachment' in filter.keys():
        if filter['has_attachment']:
            messages = messages.Restrict("@SQL=urn:schemas:httpmail:hasattachment=1")

    return list(messages)

def _keep_attachment(attachment, filter_out : tuple) -> bool:
    """Takes and attachment and if filename contains any of the strings in filter_out returns True
    """
    keep = True
    for filter in filter_out:
        if filter in attachment.FileName:
            flag=False
    return keep

def download_attachments(messages : list, folder : str, filter_out : tuple=('.jpg',), change_filename : bool=False) -> list:
    """Downloads the attachments from a collection of messages to a specified folder. If change_filename is True the
    names will be generated to have format `request2_3'. Function will overwrite files with out checking.

    Args:
        messages (list): list of message items returned by get_emails. Takes output from get_emails()
        folder (tuple, optional): folder to which attachments should be downloaded specified in 
        hierarchical format ('Inbox', 'Admin', 'Church').
        filter_out (tuple, optional): will not download attachments containing strings in tuple
        change_name: (bool): changes filename to str(uuid.uui4()) if True else keeps filename, defaults to False

    Returns:
        list: list of strings containing downloaded attachment full path and new filenames.
    """
    attachment_names = []
    for message in messages:
        for attachment in message.Attachments:
            if _keep_attachment(attachment, filter_out):
                if change_filename:
                    filename=folder + str(uuid.uuid4()) + pathlib.Path(attachment.FileName).suffix
                else:
                    filename = attachment.FileName
                attachment.SaveAsFile(filename)
                attachment_names.append(filename)
    return attachment_names

def move_emails(outlook : Type[win32.Dispatch], messages : list, folder : tuple=('Inbox')):
    """Moves messages to new folder

    Args:
        outlook (Type[win32.Dispatch]): Mail Object
        messages (list): list of message items returned by get_emails. Takes output from get_emails()
        folder (tuple, optional): tuple of strs corresponding to hierarcy of folders. Defaults to ('Inbox').

    Raises:
        FolderNotFoundException: _description_
    """

    to_mailbox = find_folder(outlook, folder=folder)

    for message in messages:
        message.move(to_mailbox)

def _check_attachments(msg, attachments):
    #Does quick check that all attachments have loaded correctly. Sometimes Onedrive syncing prevents this.
    return len(msg.Attachments) == len(attachments)


def send_email(msg: dict, attachments=None, attempt=1, max_attempts=5):
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
        attached_correctly = _check_attachments(mail, attachments)
        if not attached_correctly:
            if attempt <= max_attempts:
                time.sleep(30)
                send_email(msg, attachments=attachments, attempt=attempt+1)
            else:
                raise EmailAttachmentException()

    mail.Send()



"""
------------------------------------------------------------------------------
Helper functions
------------------------------------------------------------------------------
"""

def print_folder_names(outlook, account='ppzmis@exmail.nottingham.ac.uk'):
    """
    Quick utility method to show folder names and folder tree. Construct
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

def extract_unique_properties(messages):
    """Extract the unique values in certain fields from a collection of messages

    Args:
        messages (Type[win32.CDispatch]): Collection of messages

    Returns:
        dictionary of lists: each list contains strings pertaining to unique values
    """
    properties = {
                    'from_email':[],
                    'from_name':[],
                    'subject':[]
                }
    for message in messages:
        print(message.Subject)

    for message in messages:
        properties['from_name'].append(str(message.Sender))
        properties['from_email'].append(str(message.SenderEmailAddress))
        properties['subject'].append(str(message.Subject))

    properties['from_name'] = list(set(properties['from_name']))
    properties['from_email'] = list(set(properties['from_email']))
    properties['subject'] = list(set(properties['subject']))

    return properties


