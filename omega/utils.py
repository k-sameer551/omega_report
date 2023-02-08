import os, time
from datetime import datetime, timedelta
from pathlib import Path
import psutil
import win32com.client as win32
from win32comext.shell import shell, shellcon
# import win32gui, win32con


class Utils():
    """utitliy function"""
    def __init__(self) -> None:
        pass
    
    @classmethod
    def share_dynamic(cls, files_list: list):
        """email dynamic report"""
        Utils.close_app('outlook')
        todays_date = datetime.today() - timedelta(hours=10, minutes=30)
        outlook = win32.Dispatch('Outlook.Application')
        wsh = win32.Dispatch('Wscript.Shell')
        olmailitem=0x0
        mail_item = outlook.CreateItem(olmailitem)
        mail_item.Display()
        # signature = mail_item.HTMLBody
        mail_item.To = 'abc'
        mail_item.CC = 'abc'
        mail_item.Subject = 'Omega Dynamic Report ' + todays_date.strftime('%m-%d-%Y')  # %H:%M %p
        mail_item.HTMLBody = r"""
                Dear All,<br><br>
                Here by attached Omega Dynamic Production Report date """ \
                    + todays_date.strftime('%m-%d-%Y') + r""".<br><br>""" \
                    + r"""Thanks and Regards <br> UNET Auto Reporting""" 
        for file in files_list:
            mail_item.Attachments.Add(file)
        mail_item.Display()
        # mail_item.Send()
        wsh.AppActivate("Outlook")
        time.sleep(2)
        wsh.SendKeys("%s", 0)
        # if win32gui.FindWindowEx(win32gui.FindWindow(None, "Microsoft Outlook"), 0, "Static", "DAL=on"):
        #     window = win32gui.FindWindow(None, "Microsoft Outlook")
        #     allow_btn = win32gui.FindWindowEx(window, 0, "Button", "Allow")
        #     win32gui.SendMessage(allow_btn, win32con.BM_CLICK, 0, 0)
        #     time.sleep(1)


    @classmethod
    def get_alldetail_files_path(cls, download_folder):
        """ return list of alldetails files"""
        files = []
        for file in Path(download_folder).iterdir():
            if file.name.find('Dynamic Processed Work Item') != -1:
                files.append(os.path.join(download_folder, file.name))
        return files

    @classmethod
    def close_app(cls, app_name):
        """close app"""
        running_apps=psutil.process_iter(['pid','name']) #returns names of running processes
        found=False
        for app in running_apps:
            sys_app=app.info.get('name').split('.')[0].lower()
            if sys_app in app_name.split() or app_name in sys_app:
                pid=app.info.get('pid') #returns PID of the given app if found running
                try: #deleting the app if asked app is running.(It raises error for some windows apps)
                    app_pid = psutil.Process(pid)
                    app_pid.terminate()
                    found=True
                except: pass
            else: pass
        if not found:
            return False
        else:
            return True

    @classmethod
    def rename_file(cls, new_name, folder_path):
        """rename the file"""
        for file in Path(folder_path).iterdir():
            if file.name.startswith('Dynamic Processed'):
                Path.rename(file.name, new_name + " - " + file.name)
                break
        return True

    @classmethod
    def delete_file(cls, folder_path):
        """delete previously downloaded file"""
        for file in Path(folder_path).iterdir():
            if file.name.find('Dynamic Processed') != -1:
                file.unlink()
        return True

    @classmethod
    def get_download_path(cls):
        """get the path for downloading omega file"""
        return shell.SHGetFolderPath(0,shellcon.CSIDL_PERSONAL, None, 0)
    
    @classmethod
    def set_working_directory(cls, folder_path):
        """set working directory path"""
        return os.chdir(folder_path)