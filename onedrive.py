import sys
sys.path.append('..')

import win32api

def file_locally_available(path : str, filename : str):
    """Check file is available on local machine. Used to prevent issues due to
    Onedrive syncing.

    Args:
        path (str) : filename to path
        filename (str): filename

    Returns:
        bool: True if file available locally, otherwise False
    """
    attr = win32api.GetFileAttributes(path + filename)
    print(attr)
    print('end')
    #return available

if __name__ == '__main__':
    path = 'C:/Users/ppzmis/OneDrive - The University of Nottingham/Documents/Programming/emails/'
    filename = 'setup.py'
    file_locally_available(path, filename)