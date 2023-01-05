

class FolderNotFoundException(Exception):
    def __init__(self, e):
        super().__init__(e)
        print('Folder not found. Check with ')

class EmailAttachmentException(Exception):
    def __init__(self, e):
        super().__init__(e)
        print('Failure to send email with attachment')