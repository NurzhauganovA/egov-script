import pandas as pd
import first_part_send_order
import fourth_part


CAPITALIZED_NAMES = [
    'AKT_ORSK',
    'ATR_ALMALY2', 'ATR_BEKTAS', 'ATR_ERKIN', 'ATR_ILICHA',
    'ATR_INDERBOR', 'ATR_INFECTION', 'ATR_KOKARNA', 'ATR_KOKINBOR',
    'ATR_OBSHAGA', 'ATR_SHALKYMA', 'ATR_TCON',
    'TGZ_AVTODOR', 'TGZ_BEYBARS',
    'OSK_AKIMOVKA', 'OSK_FLY', 'OSK_HAVANA', 'OSK_HUNTER',
    'OSK_MILLER', 'OSK_PROM', 'OSK_RENESANS', 'OSK_RODNICHOK',
    'OSK_SAMAL', 'OSK_SAMOCVET', 'OSK_SHERBAKOV', 'OSK_VITYAZ',
    'KAR_AIRYQPRM', 'KAR_AKSHIYPRM', 'KAR_AKSHOKYPRM', 'KAR_KOYANDYPRM',
    'KAR_SARYOBALYPRM', 'KAR_TALDYPRM',
    'KOS_DIAMOND', 'KOS_MASSIV', 'KOS_PROMZONA',
    'PTR_TATAR'
]


ALMATY_REGIONS = [
    'Алматы',
    'Алматинская область',
    'Жетысуская область'
]


def change_region(region):
    if region in ALMATY_REGIONS:
        region = 'Алматинская, Жетысуская область и Алмата'
        return region
    return region


def change_name(name):
    if name in CAPITALIZED_NAMES:
        return name

    prefix = name[:3].upper()
    suffix = name[4:].capitalize()
    result = prefix + '_' + suffix
    return result


def generate_path(region, name) -> str:
    name = change_name(name)
    region = change_region(region)

    path = fr'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\{region}\{name}\1 этап'
    return path


def generate_file_path(region, name) -> str:
    path = fr'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\{region}\{name}\АПЗ'
    return path


df = pd.read_excel('test.xlsx', skiprows=1)

COLUMNS_TO_CHECK = [
    'Землеустроительный проект',
    'АПЗ'
]

cell_to_check = 'загрузить'

filtered_df = []

for i in COLUMNS_TO_CHECK:
    filtered_df.append(df[df[i] == cell_to_check])


for dataframe in filtered_df:
    for i in dataframe.index:
        if i in dataframe[dataframe[COLUMNS_TO_CHECK[0]] == cell_to_check].index:
            path = generate_path(
                dataframe['Регион'][i],
                dataframe['Название'][i]
            )
            first_part_send_order.main(path)
        elif i in dataframe[dataframe[COLUMNS_TO_CHECK[1]] == cell_to_check].index:
            path = generate_file_path(
                dataframe['Регион'][i],
                dataframe['Название'][i]
            )
            fourth_part.main(path)

        print(path)
