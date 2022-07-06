import pandas as pd
import matplotlib.pyplot as plt
import lasio
import os

rewrite = True


def log_plot(logs):
    if logs.index.names[0] != 'DEPT':
        logs.index.names = ['DEPT']

    logs = logs.reset_index()

    logs = logs.sort_values(by='DEPT')
    top = logs.DEPT.min()
    bot = logs.DEPT.max()

    f, ax = plt.subplots(nrows=1, ncols=len(logs.columns) - 1, figsize=(14, 8))

    cols = list(logs.columns)
    cols.remove('DEPT')

    for i in range(len(ax)):
        ax[i].plot(logs[cols[i]], logs.DEPT, color='green')
        ax[i].set_ylim(top, bot)
        ax[i].invert_yaxis()
        ax[i].grid()
        ax[i].set_xlabel(cols[i])
        ax[i].set_xlim(logs[cols[i]].min(), logs[cols[i]].max())

    for i in range(1, len(ax)):
        ax[i].set_yticklabels([])

    f.suptitle('CURVES', fontsize=14, y=0.94)


def merge_two_same_mnemonic(df, mnem, mnem_x, mnem_y):
    df[mnem_x].fillna(df[mnem_y], inplace=True)
    df.drop(columns=[mnem_y], inplace=True)
    df.rename(columns={mnem_x: mnem}, inplace=True)
    return df


def marry_lases(df11, df22):
    global rewrite
    df1 = df11.copy()
    df2 = df22.copy()

    df1.dropna(axis=1, how='all', inplace=True)
    df2.dropna(axis=1, how='all', inplace=True)

    if df1.index.names[0] != 'DEPT':
        df1.index.names = ['DEPT']
    if df2.index.names[0] != 'DEPT':
        df2.index.names = ['DEPT']

    df1.reset_index(inplace=True)
    df2.reset_index(inplace=True)

    df = pd.merge(df1, df2, on='DEPT', how='outer', suffixes=['_x', '_y'])

    columns_intercept = list(set(df1.columns) & set(df2.columns))
    columns_intercept.remove('DEPT')

    # print('общие мнемоники: ', columns_intercept)
    # if list(set(df1.DEPT) & set(df2.DEPT)):
    #     print('границы пересечения пог глубине: ', min(list(set(df1.DEPT) & set(df2.DEPT))),
    #           max(list(set(df1.DEPT) & set(df2.DEPT))), '\n')

    ### Использование mnem_x, mnem_y
    if rewrite is True:
        for col in columns_intercept:
            df = merge_two_same_mnemonic(df, col, col + '_y', col + '_x')
    else:
        for col in columns_intercept:
            df = merge_two_same_mnemonic(df, col, col + '_x', col + '_y')

    step1 = df1.DEPT[1] - df1.DEPT[0]
    step2 = df2.DEPT[1] - df2.DEPT[0]

    ### Как реализована работа с шагом, в чем суть?
    if step1 < step2:
        for i in range(len(df)):
            if df.DEPT[i] * 10 % (step2 * 10) != 0:
                df.drop([i], inplace=True)
    elif step1 > step2:
        for i in range(len(df)):
            if df.DEPT[i] * 10 % (step1 * 10) != 0:
                df.drop([i], inplace=True)

    df.sort_values(by=['DEPT'], inplace=True)

    df = df.set_index('DEPT')

    return df


def get_header_data(file_path, unique_curves):
    las = lasio.read(file_path, encoding="windows-1251")
    # las = lasio.read(file_path)
    for curve in las.curves:
        if curve.mnemonic not in unique_curves:
            unique_curves[curve.mnemonic] = curve
        else:
            pass
    return unique_curves


def get_file_data(path_to_folder):
    unique_curves = {}
    for root, folders, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".las"):
                print(file)
                file_path = os.path.join(root, file)
                current_las = lasio.read(file_path)
                current_df = current_las.df()
                try:
                    las_data = marry_lases(las_data, current_df)
                except NameError:
                    las_data = current_df.copy()

                unique_curves = get_header_data(file_path, unique_curves)
    las_data.fillna(-999.25, inplace=True)
    return las_data, unique_curves


def form_las_file(las_data, unique_curves):
    las_file = lasio.LASFile()
    las_file.add_curve(mnemonic="MD",
                       data=las_data.index.values,
                       unit="Meters",
                       descr="Mesured depth")
    for curve in las_data:
        curve_name = las_data[curve].name
        curve_data = las_data[curve].values
        curve_unit = unique_curves[curve_name].unit
        curve_description = unique_curves[curve_name].descr
        las_file.add_curve(mnemonic=curve_name,
                           data=curve_data,
                           unit=curve_unit,
                           descr=curve_description)
    return las_file


def process_folder(folder_path, result_link):
    las_data, unique_curves = get_file_data(folder_path)
    las_file = form_las_file(las_data, unique_curves)
    las_file.write(result_link, version=2.0, fmt='%.4f')


if __name__ == '__main__':
    folder_path = r''
    result_link = r""
    process_folder(folder_path, result_link)
