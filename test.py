from emails import auto_email
from pydates.pydates import relative_datetime, now


outlook = auto_email.open_outlook()

filter = {'start': relative_datetime(now(),delta_day=-1000),
              'stop':relative_datetime(now(),delta_day=-1),
            }

#msgs = auto_email.get_emails(outlook, ('Inbox','DLO','coursework_extensions'), filter=filter)
#print(auto_email.extract_unique_properties(msgs))
print(auto_email.find_sender_emails(outlook, ('Inbox','DLO','coursework_extensions')))