import os
import shutil

def file_preprocess(source_folder, destination_folder):
    # 将爬取的excel表复制一份到merge文件夹下面
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(source_folder):
        if filename.endswith(".xlsx"):
            source_path = os.path.join(source_folder, filename).replace(os.sep, '/')
            destination_path = os.path.join(destination_folder, filename).replace(os.sep, '/')
            shutil.copy(source_path, destination_path)

    # 将merge文件夹下面的所有excel文件统一命名
    index = 0
    for filename in os.listdir(destination_folder):
        if filename.endswith(".xlsx"):
            index += 1
            file_path = os.path.join(destination_folder, filename).replace(os.sep, '/')
            new_filename = f"data ({index}).xlsx"
            new_file_path = os.path.join(destination_folder, new_filename).replace(os.sep, '/')
            shutil.move(file_path, new_file_path)