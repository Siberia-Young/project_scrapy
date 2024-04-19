import os
import shutil

def file_select(destination_folder, outcome_folder, target_list):
    # 将目标excel表复制一份到outcome文件夹下面
    if not os.path.exists(outcome_folder):
        os.makedirs(outcome_folder)
    file_list = []
    for filename in os.listdir(destination_folder):
        # 找到以merge开头的excel文件
        if filename.endswith(".xlsx") and filename.startswith("merge"):
            source_file = os.path.join(
                destination_folder, filename).replace(os.sep, '/')
            file_list.append(source_file)
            
    num = 1
    for filename in target_list:
        for path in file_list:
            if path.split('/')[-1] == filename:
                destination_file = os.path.join(outcome_folder, '文件'+str(num)+'.xlsx').replace(os.sep, '/')
                shutil.copy(path, destination_file)
                num += 1