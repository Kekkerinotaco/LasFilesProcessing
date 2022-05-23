import welly
import os
import shutil
import openpyxl

def main():
    # Ссылка на папку с las файлами для обработки
    input_folder = r""
    # Ссылка на папку, в которой будет необходимо сохранить результат
    result_folder = r""
    well_dictionary, errors_dictionary = determine_wells(input_folder)
    manage_files(well_dictionary, result_folder, copy=True)
    write_errors(errors_dictionary, result_folder)



def determine_wells(link):
    """ Функция получает на вход ссылку на папку с las файлами, на выходе отдает словарь формата
    Скважина: ссылка на файлы относящиеся к ней


    :param link: Ссылка на папку содержащую las файлы
    :return: словарь Формата Скважина: ссылка на файлы относящиеся к ней
    """
    well_dictionary = {}
    errors_dictionary = {}
    for root, folders, files in os.walk(link):
        for file in files:
            print(file)
            file_path = os.path.join(root, file)
            try:
                well = welly.Well.from_las(file_path, encoding="windows-1251")
                if well.name not in well_dictionary.keys():
                    well_dictionary[well.name] = []
                    well_dictionary[well.name].append(file_path)
                else:
                    well_dictionary[well.name].append(file_path)
            except Exception as e:
                errors_dictionary[file_path] = str(e)
    return well_dictionary, errors_dictionary


def manage_files(well_dictionary, result_folder, copy=True):
    """ Функция получает на вход словарь из функции determine_wells, создает папки с названием ключей, копирует файлы
    находящиеся в значениях


    :param well_dictionary: Словарь [название : ссылка на файлы, относящиеся к ней] из функции determine_wells
    :param result_folder: Словарь [ссылка на файл : ошибка] из функции determine_wells
    :param copy:
    :return:
    """

    for well_name, list_of_links in well_dictionary.items():
        folder = os.path.join(result_folder, str(well_name))
        try:
            os.mkdir(folder)
        except:
            print("Problem with well: {}".format(well_name))
        if copy is True:
            for link in list_of_links:
                shutil.copy(link, folder)
        if copy is False:
            for link in list_of_links:
                shutil.move(link, folder)


def write_errors(errors_dictionary, result_folder):
    """ Функция получает на вход словарь с ошибками из функции determine_wells, создает папки с названием ключей,
    копирует файл находящиеся в значениях


    :param errors_dictionary: Словарь [Ссылка на файл: ошибка] из функции determine_wells
    :param result_folder: Ссылка на папку, в которой будет создан Excel-файл, с ошибками
    :return:
    """
    file_path = os.path.join(result_folder, "Errors.xlsx")
    errors_workbook = openpyxl.Workbook()
    worksheet = errors_workbook.worksheets[0]
    row_to_write = 2
    worksheet["A1"].value = "Filepath"
    worksheet["B1"].value = "Error"
    for file, error in errors_dictionary.items():
        worksheet["A{}".format(row_to_write)].value = file
        worksheet["B{}".format(row_to_write)].value = error
        row_to_write += 1
    errors_workbook.save(file_path)


if __name__ == '__main__':
    main()

