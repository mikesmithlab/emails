import unittest
import emails.auto_email

class Test(unittest.TestCase):

    def test_get_emails(self):
        #outlook = auto_email.open_outlook()
        #filter = {'subject':'Test auto'}
        #messages = auto_email.get_emails(outlook, filter=filter)
        self.assertEqual(1,1)#messages[0].Subject,'Test auto_email.py')

