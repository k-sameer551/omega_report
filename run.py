from omega.omega import Omega
from omega.utils import Utils

with Omega(teardown=False) as bot:
    download_folder = Utils.get_download_path()
    Utils.delete_file(download_folder)
    Utils.set_working_directory(download_folder)
    bot.land_web_page()
    bot.download_report('CSS CRT', 'CRT', download_folder, 60)
    bot.download_report('CSS UNET Review', 'Review', download_folder, 60)
    bot.download_report('CSS UNET Rework', 'Rework', download_folder, 30)
    files_list = Utils.get_alldetail_files_path(download_folder)
    Utils.share_dynamic(files_list)
    bot.quit()