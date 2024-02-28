import argparse
import pandas as pd
import chardet
import sys
import re
import datetime
# D:\kozyakov\lab1\rem.xls D:\kozyakov\lab1\источник1.txt D:\kozyakov\lab1\vl.csv

pd.set_option('display.width', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)


# Считываем все данные из всех файлов
def load_data(file_paths):
    dfs = []

    for file_path in file_paths:
        file_extension = file_path.split(".")[-1].lower()

        if file_extension in ["xls", "xlsx"]:
            data = pd.read_excel(file_path, header=0)
            dfs.append(data.astype(str))
        elif file_extension == "csv":
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read())
            data = pd.read_csv(file_path, delimiter=';', encoding=result['encoding'], header=None).astype(str)
            dfs.append(data)
        elif file_extension == "txt":
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read())
            data = pd.read_csv(file_path, delimiter='\t', encoding=result['encoding'], header=None).astype(str)
            dfs.append(data)
        else:
            print(f"Unsupported file format for file: {file_path}")
            continue
    return dfs


def transform_data(data):
    # Конвертируем значения в колонка в строковые
    data.columns = data.columns.astype(str)

    # Поиск всех столбцов, в названии которых есть слово 'цена'
    price_columns = [col for col in data.columns if re.search(r'(цена|стоимость)', str(col), re.IGNORECASE)]

    # Очистка всех значений кроме числовых в найденных столбцах цены
    for col in price_columns:
        if col in data.columns:
            data[col] = data[col].apply(lambda x: re.sub(r'\D', '', x) if isinstance(x, str) else x)

    name_columns = [col for col in data.columns if re.search(r'(ФИО|мастер|клиент)', str(col), re.IGNORECASE)]

    # Преобразование данных ФИО в формат Фамилия И. или Фамилия И. О. если они уже сокращены
    for column in name_columns:
        for index, value in data[column].items():
            parts = value.split()
            if len(parts) == 2:
                # Если ФИО записано в формате "Фамилия И.О."
                surname = parts[0]
                initials = parts[1]
                if len(initials) == 4 and initials[1] == '.':
                    # Добавляем пробел между инициалами
                    processed_name = f"{surname} {initials[0]}. {initials[2]}."
                    data.at[index, column] = processed_name

    # Преобразование данных ФИО в формат Фамилия И. или Фамилия И. О. если они идут в полном формате "Фамилия Имя Отчество"
    for column in data.columns:
        for index, value in data[column].items():
            parts = value.split()
            if len(parts) >= 2:
                # Проверка, что первые буквы слов написаны с заглавной буквы
                if all(part.istitle() for part in parts):
                    last_name = parts[0]
                    first_name_initial = parts[1][0]
                    middle_name_initial = parts[2][0] if len(parts) > 2 else ''
                    # Проверка, что в каждой части имени содержатся только буквы
                    if all(part.isalpha() for part in parts):
                        processed_name = f"{last_name} {first_name_initial}. {middle_name_initial}." if middle_name_initial else f"{last_name} {first_name_initial}."
                        data.at[index, column] = processed_name

    # Преобразование даты в формат 'день месяц год'
    data_columns = [colu for colu in data.columns if re.search(r'дата', str(colu), re.IGNORECASE)]
    for colu in data_columns:
        for index, value in data[colu].items():
            if re.match(r'^\d+$', value):
                days = int(value)
                date = pd.to_datetime('1900-01-01') + pd.to_timedelta(days, unit='D')
                data.at[index, colu] = date.strftime('%d.%m.%Y')

    #Изменяем формат записи паспорта с хххх/ хххххх на хххх хххххх, и если формат паспорта не совпадает (в номере не 6 значений, а более, убираем последние значния чтобы их осталось 6)
    for column in data.columns:
        for index, value in data[column].items():
            parts = value.split('/')
            if len(parts) == 2:
                series = parts[0].strip()
                number = parts[1].strip()
                # Проверяем последнюю цифру в номере
                if len(number) > 6:
                    number = number[:6]
                processed_value = f"{series} {number}"
                data.at[index, column] = processed_value

    return data


#Сохранение итоговых табиц в EXCEL файл по указаннному пути
def save_all_tables_to_excel(tables, folder_path):
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_path = folder_path.rstrip('/') + f'/Итог_{current_datetime}.xlsx'
    with pd.ExcelWriter(file_path) as writer:
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name, index=False)
    print(f'Все таблицы сохранены в файл: "{file_path}"')
    print_fireworks()


#Красоты ради=)
def print_fireworks():
    fireworks_side = r"""
         .   *  *  " .  *  .  *  *  * .   *  *  " .  *  .  *  * "  .   *  *  " .  *  .  *  * . *  *  " .  *  .  *  *  * 
          * . " *   . *   .  * . "   .  * . " *   . *   .  * . "  .   * . " *   . *   .  * . " * . " *   . *   .  * . " 
       *  . "   .  *  *  .  "  .  *  *  . "   .  *  *  .  "  .  *  *  . "   .  *  *  .  "  .  *  *  . "   .  *  *  .  " 
          .  * " .  *  .  " .  *  .  "   .  * " .  *  .  " .  *  .  .   .  * " .  *  .  " .  *  .  .  * " .  *  .  " .  
          *   . "  * . "  * . "   *  .   *   . "  * . "  * . "   *  "   *   . "  * . "  * . "   *   *   . "  * . "  * . 
         .   *  " .  *  .  "  *  .    "   .   *  " .  *  .  "  *  .  * "   .   *  " .  *  .  "  .  .   *  " .  *  .  "  
           """

    print(fireworks_side)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        #Указываем в консоль путь к файлам которые нам необходимо обработать
        print("Укажите путь к необходимым файлам.")
        file_paths = input("Пути к файлам (через пробел): ").split()
    else:
        parser = argparse.ArgumentParser(description='Process some files.')
        parser.add_argument('files', metavar='file', type=str, nargs='+', help='List of files to process')
        args = parser.parse_args()
        file_paths = args.files

    data_frames = load_data(file_paths)
    transformed_dfs = [transform_data(df) for df in data_frames]

    #объединяем таблицы 2 и 3 (источник1.txt и vl.csv)
    merged_data = (pd.merge(transformed_dfs[1], transformed_dfs[2], left_on='4', right_on='1', how='inner')
                   .drop(columns=['2_y', '1_y'])
                   .rename(columns={'0_x': 'Клиент', '1_x': 'Фирма-производитель', '2_x': 'Марка',
                                    '3_x': 'Год', '4': 'Паспорт', '0_y': 'Права', '3_y': 'WIN'}))

    # Объединяем созданную таблицу (с данными из источник1.txt и vl.csv) и rem.xls
    merged_end = (pd.merge(transformed_dfs[0], merged_data, left_on='ВИН', right_on='WIN', how='inner')
                    .drop(columns=['WIN'])
                    .rename(columns={'ВИН': 'WIN', 'кол-часов': 'Кол-во часов', 'Коээффициент мастера': 'Коэффициент мастера'}))

    # Создание таблиц по типу снежинка
    merged_firm = merged_end[['Фирма-производитель']].copy().drop_duplicates()
    merged_firm.insert(0, 'id_firm', range(len(merged_firm)))
    print('Фирма производитель')
    print(merged_firm)
    print('- ' * 50)

    #Создание таблицы Фирма-производитель
    merged_brand1 = pd.merge(merged_end[['Марка', 'Фирма-производитель']].copy(),
                            merged_firm[['id_firm', 'Фирма-производитель']].copy()).drop_duplicates()
    merged_brand1.insert(0, 'id_brand', range(len(merged_brand1)))

    merged_brand = merged_brand1.drop(columns=['Фирма-производитель'])
    print('Марка')
    print(merged_brand)
    print('- ' * 50)

    #Создание таблицы Автомобиль
    merged_auto = pd.merge(merged_end[['Год', 'WIN', 'Фирма-производитель']].copy(), merged_brand1[['id_brand', 'id_firm', 'Фирма-производитель']].copy())\
                            .drop_duplicates()
    if 'Фирма-производитель' in merged_auto.columns:
        merged_auto = merged_auto.drop('Фирма-производитель', axis=1)
    merged_auto.insert(0, 'id_auto', range(len(merged_auto)))
    print("Автомобиль")
    print(merged_auto)
    print('- ' * 50)

    #Создание таблицы Потребитель
    merged_client = pd.merge(merged_end[['Клиент', 'Паспорт']], merged_auto[['id_auto']], left_index=True, right_index=True)\
                            .drop_duplicates()
    merged_client.insert(0, 'id_client', range(len(merged_client)))
    print('Потребитель')
    print(merged_client)

    #Создание таблицы Операция
    merged_operation = merged_end[['Операция']].drop_duplicates()
    merged_operation.insert(0, 'id_operation', range(len(merged_operation)))
    print("Операция")
    print(merged_operation)
    print('- ' * 50)

    #Создание таблицы Мастер
    merged_master = merged_end[['Мастер']]\
                            .drop_duplicates()
    merged_master.insert(0, 'id_master', range(len(merged_master)))
    print("Мастер")
    print(merged_master)
    print('- ' * 50)

    #Создание таблицы Факты
    merged_facts_1 = pd.merge(merged_end[['Дата', 'WIN', 'Кол-во часов', 'Коэффициент мастера', 'Цена ремонта', 'Мастер']], merged_master[['id_master', 'Мастер']])\
                            .drop_duplicates()

    merged_facts = pd.merge(merged_facts_1, merged_operation[['id_operation', 'Операция']], left_index=True, right_index=True)\
                            .drop_duplicates()\
                            .drop(columns = ['Мастер', 'Операция'])
    merged_facts.insert(0, 'id_facts', range(len(merged_facts)))
    print("Факты")
    print(merged_facts)
    print('- ' * 50)

    tables_to_save = {
        'Фирма производитель': merged_firm,
        'Марка': merged_brand,
        'Автомобиль': merged_auto,
        'Потребитель': merged_client,
        'Операция': merged_operation,
        'Мастер': merged_master,
        'Факты': merged_facts
    }

    folder_path = input("Введите путь для сохранения итогового файла: ")
    save_all_tables_to_excel(tables_to_save, folder_path)


