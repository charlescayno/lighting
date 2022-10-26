import os
import shutil

def path_maker(file_path: str):
    """
        file_path: Enter the file path you want to create

        returns the string value of new path created
    """

    folder_list = file_path.split('/')

    new_path = ' '
    for i in folder_list:
        
        if new_path == ' ':
            path = f'{i}/'
        else: 
            path = new_path + f'{i}/'

        if not os.path.exists(path):
            os.mkdir(path)
            
        new_path = path
    return new_path


def remove_file(file_path: str):
    """
        file_path: Enter the file path you want to delete

        returns the string value of new path created
    """

    if os.path.exists(file_path): os.remove(file_path)


def move_file(source_path: str, destination_path: str):
    """
        source_path : path/file
        destination_path : path/file
    """
    source = f'{source_path}'
    destination = f'{destination_path}'
    remove_file(destination)
    shutil.move(source, destination)