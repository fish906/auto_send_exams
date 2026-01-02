import os
from datetime import datetime
from time import sleep


def get_file_path(file_directory):
    sub_directory = [f for f in os.listdir(file_directory) if f.endswith('_file')]
    file_paths_list = []

    for directory in sub_directory:
        full_path = os.path.join(file_directory, directory)
        sub_path = os.listdir(full_path)
        for item in sub_path:
            file_paths_list.append(os.path.join(full_path, item))

    return file_paths_list


def move_files(file_path):
    os.makedirs("exams", exist_ok=True)

    for old_path in file_paths:
        filename = os.path.basename(old_path)
        new_path = os.path.join("exams", filename)

        try:
            os.replace(old_path, new_path)
            print(f'Moved {old_path} to {new_path}')

        except Exception as e:
            print(e)

if __name__ == '__main__':
    start_time = datetime.now()
    directory = "attachments"
    file_paths = (get_file_path(directory))

    move_files(file_paths)
    print(datetime.now() - start_time)