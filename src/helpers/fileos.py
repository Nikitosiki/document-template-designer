import os

def delete_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)

def delete_file(is_delete, file_path):
    if is_delete and os.path.exists(file_path):
        os.remove(file_path)