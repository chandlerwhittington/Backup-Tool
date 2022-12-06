import os
import getpass
import datetime
import shutil
import subprocess


backup_location = ''
today = datetime.datetime.now()
backup_location = f'Laptop Backup {today.month}.{today.day}.{today.year}'
username = getpass.getuser()
dir_with_onedrive = os.path.join("C:\\", "Users", username, "OneDrive - Guaranteed Rate Inc")
dir_without_onedrive = os.path.join("C:\\", "Users", username)
dir_for_microsoft = os.path.join(dir_with_onedrive, 'AppData', 'Roaming', 'Microsoft')


#use this to see if a dir exists and create a backup folder
if os.path.exists(dir_with_onedrive) or os.path.isdir(dir_with_onedrive):
    dir_for_backup = os.path.join(dir_with_onedrive, backup_location)

    if not os.path.exists(dir_for_backup):
        os.mkdir(dir_for_backup)

        print(f'The folder {dir_for_backup} was created')
    else:
        print(f'The folder {dir_for_backup} already exists')
    print(f'The {dir_with_onedrive} is a directory')
else:
    dir_for_backup = os.path.join(dir_without_onedrive, backup_location)
    if not os.path.exists(dir_for_backup):
        os.mkdir(dir_for_backup)
        print(f'The folder {dir_for_backup} was created')
    else:
        print(f'The folder {dir_for_backup} already exists')
    print(f'The {dir_with_onedrive} is NOT a directory')


#create the subfolders of the backup folder
backup_folders =['Edge', 'FireFox', 'Google Chrome', 'Microsoft', 'PrinterExport', 'Wifi Profile']
try:
    for folders in backup_folders:
        create_folder = os.path.join(dir_for_backup, folders)
        if not os.path.exists(create_folder):
            os.makedirs(create_folder)
        # print(create_folder)
except FileExistsError:
    print('file exists')


#backs up the Microsoft data
dirs_for_microsoft =['Signatures','Speech', 'Stationery', 'Templates', 'Sticky Notes']
try:
    for dirs in dirs_for_microsoft:
        backup_source_ms = os.path.join(dir_without_onedrive, 'AppData', 'Roaming', 'Microsoft', dirs)
        dest_dir = os.path.join(dir_for_backup, 'Microsoft', dirs)
        files = os.listdir(backup_source_ms)
        shutil.copytree(backup_source_ms, dest_dir)
except FileNotFoundError:
    print(f'the file location {backup_source_ms} doesnt exist')
except FileExistsError:
    print(f'the file {backup_source_ms} exists')

#backup edge firefox and chrome bookmarks

chrome_bm = os.path.join(dir_without_onedrive, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
edge_bm = os.path.join(dir_without_onedrive, 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data', 'Default', 'Bookmarks')
chrome_folder = os.path.join(dir_for_backup, 'Google Chrome')
edge_folder = os.path.join(dir_for_backup, 'Edge')

shutil.copy(chrome_bm, chrome_folder)
shutil.copy(edge_bm, edge_folder)


#printExport
try:
    print_export = 'C:\\Windows\\System32\\spool\\tools\\PrintBrm.exe'
    printerBackupName = f"printers_{today.month}.{today.day}.{today.year}.printerExport"
    dir_for_printer_backup = os.path.join(dir_without_onedrive, 'PrinterExport')
    if not os.path.exists(dir_for_printer_backup):
        os.makedirs(dir_for_printer_backup)

    printer_backup_location = os.path.join(dir_for_printer_backup, printerBackupName)
    print_backup = f'start {print_export} -B -f {printer_backup_location}'
    os.system(print_backup)
    #print_copy_location = os.path.join(dir_for_printer_backup)
    print_copy_to = os.path.join(dir_for_backup, 'PrinterExport')

    shutil.copy(printer_backup_location, print_copy_to)
except FileExistsError:
    print("location exists")

WIFI_Backup = f'WIFI Backup {today.month}.{today.day}.{today.year}.txt'
wifi_pre_folder = os.path.join(dir_for_backup, 'Wifi Profile')
wifi_folder = os.path.join("C:\\", "Users", username, "OneDrive - Guaranteed Rate Inc", backup_location, wifi_pre_folder, WIFI_Backup)

data = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8').split('\n')
profiles = [i.split(":")[1][1:-1] for i in data if "All User Profile" in i]
for i in profiles:
    results = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', i, 'key=clear']).decode('utf-8').split('\n')
    results = [b.split(":")[1][1:-1] for b in results if "Key Content" in b]
    try:
        with open (wifi_folder, 'a') as f:
            f.writelines("{:<30}|  {:<}".format(i, results[0]) + '\n')

    except IndexError:
        with open('WIFI Backup.txt', 'a') as f:
            f.write("{:<30}|  {:<}".format(i, "") + '\n')