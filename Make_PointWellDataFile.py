import lasio
import os
import openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

result_DF = pd.DataFrame()


def main():
    # Ссылка на папку с файлами для обработки
    link = r""
    # Получение возможных типов объектов на основании названия папок, используется далее при создании DataFrame
    folder_names_list = get_folders_names(link)
    result_link = os.path.join(link, "Point_well_Data.xlsx")
    for folder_name in folder_names_list:
        # Обработка папок по типам объектов
        folder_path = os.path.join(link, folder_name)
        process_folder(folder_path)
    export_to_excel(result_link)


def get_folders_names(link):
    """ Функция формирует список лист из названия


    :param link:
    :return:
    """
    list_of_object_types = [folder for folder in os.listdir(link) if os.path.isdir(os.path.join(link, folder))]
    # list_of_well_types = []
    # folder_items = os.listdir(link)
    # for item in folder_items:
    #     item_path = os.path.join(link, item)
    #     if os.path.isdir(item_path):
    #         list_of_well_types.append(item)
    return list_of_object_types


def process_folder(folder_path):
    """ Функция получает на вход ссылку на папку с файлами для обработки, и формирует DataFrame, содержащий необходимые
    данные. Внутри функции необходимо уточнять, что должно содержаться в DataFrame (помечено хештегами)


    :param folder_path: Путь на папку для обработки
    :return:
    """
    global result_DF
    folder_name = os.path.basename(folder_path)
    for root, folders, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            las = lasio.read(file_path, encoding="windows-1251")
            well_name = las.well.WELL.value
            las_data = las.df()
            las_data["Well"] = well_name
            # las_data["Zaboy"] = las.well.STOP.value
            las_data["Object_type"] = folder_name
            # las_data["File_name"] = file
            # las_data["File_folder"] = root
            # las_data["File_path"] = file_path
            result_DF = pd.concat([result_DF, las_data])
            print(result_DF.shape)


def export_to_excel(result_link):
    """ Экспорт Dataframe с данными в Excel
    """
    result_workbook = openpyxl.Workbook()
    result_worksheet = result_workbook.worksheets[0]
    for row in dataframe_to_rows(result_DF, index=True, header=True):
        result_worksheet.append(row)
    result_workbook.save(result_link)
    print(result_link)


if __name__ == '__main__':
    main()
