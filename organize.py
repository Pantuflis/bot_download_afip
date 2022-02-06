import os
import pandas as pd

def get_desktop_path():
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    return desktop_path 

def get_clients():
    df = pd.read_excel('files/users.xlsx')
    clients = df['CLIENT'].to_list()
    return clients

def create_paths():
    desktop_path = get_desktop_path()
    if os.path.exists(desktop_path + '\\Facturas'):
        download_path = desktop_path + '\\Facturas'
    else:
        os.makedirs(desktop_path + '\\Facturas')
        download_path = desktop_path + '\\Facturas'
    clients = get_clients()
    for client in clients:
        if not os.path.exists(download_path + f'\\{client}'):
            os.makedirs(download_path + f'\\{client}')