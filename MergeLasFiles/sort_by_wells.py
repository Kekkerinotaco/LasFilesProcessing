import os
import shutil


def main():
    # Link to a folder with files to sort
    link = r""
    # Link to a folder where the result will be stored
    result_folder = r""
    wells_dict = get_wells_dict(link)
    sort_files(result_folder, wells_dict, copy=True)


def get_wellname(file):
    index = file.find("_")
    wellname = file[:index]
    return wellname


def get_wells_dict(link):
    wells_dict = {}
    for root, folders, files in os.walk(link):
        for file in files:
            file_path = os.path.join(root, file)
            wellname = get_wellname(file)
            if wellname in wells_dict:
                wells_dict[wellname].append(file_path)
            else:
                wells_dict[wellname] = []
                wells_dict[wellname].append(file_path)
    return wells_dict


def sort_files(result_folder, wells_dict, copy=True):
    for wellname, file_paths in wells_dict.items():
        well_dir_path = os.path.join(result_folder, wellname)
        try:
            os.mkdir(well_dir_path)
        except FileExistsError:
            pass
        for file_path in file_paths:
            if copy is True:
                shutil.copy(file_path, well_dir_path)
            else:
                shutil.move(file_path, result_folder)


if __name__ == '__main__':
    main()
