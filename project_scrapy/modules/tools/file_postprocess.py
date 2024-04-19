import os
import shutil
import datetime

def file_postprocess(source_folder):
    current_time = datetime.datetime.now()
    time_string = current_time.strftime("%Y-%m-%d")
    destination_folder = os.path.join(source_folder, time_string).replace(os.sep, '/')
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(source_folder):
        if filename.endswith(".xlsx") or filename == 'merge' or filename == 'json':
            source_path = os.path.join(source_folder, filename).replace(os.sep, '/')
            destination_path = os.path.join(destination_folder, filename).replace(os.sep, '/')
            shutil.move(source_path, destination_path)