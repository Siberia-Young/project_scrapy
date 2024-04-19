import os
import shutil
import datetime

def file_package(platform_name, source_folder, outcome_folder, final_folder):
    platform = {'jd': '京东', 'tb': '淘宝', 'pdd': '拼多多批发', '1688': '1688'}
    for filename in os.listdir(source_folder):
        if filename.startswith("文件与流程关系图"):
            from_path = os.path.join(source_folder, filename).replace(os.sep, '/')
            to_path = os.path.join(outcome_folder, filename).replace(os.sep, '/')
            shutil.copy(from_path, to_path)

    current_time = datetime.datetime.now()
    folder_string = current_time.strftime(f"{platform[platform_name]}_%Y-%m-%d_%H%M%S")
    time_date_string = current_time.strftime("%Y-%m-%d")
    aim_folder = os.path.join(final_folder, time_date_string).replace(os.sep, '/')
    if not os.path.exists(aim_folder):
        os.makedirs(aim_folder)
    shutil.copytree(outcome_folder, os.path.join(aim_folder, folder_string).replace(os.sep, '/'))