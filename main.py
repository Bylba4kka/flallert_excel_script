import openpyxl
from openpyxl.utils import column_index_from_string

def replace_words(replacements: dict, phrases: list):
    # Новый список для хранения результата
    result = []

    # Итерируем по списку фраз
    for phrase in phrases:
        replaced = False  # Флаг, чтобы проверить, была ли замена
        
        # Проверяем каждый ключ из словаря
        for key, values in replacements.items():
            if key in phrase:
                # Добавляем оригинальную фразу в результат
                result.append(phrase.replace("(", "").replace(")", ""))
                
                # Если ключ найден в фразе, делаем замены и добавляем в результат
                for value in values:
                    result.append(phrase.replace(key, value))
                replaced = True
        
        if not replaced:
            # Если замена не производилась, добавляем исходную фразу
            result.append(phrase)

    # Выводим результат
    return result


FILENAME = "file.xlsx"
# Открываем файл
print("Открываем файл..")
workbook = openpyxl.load_workbook(FILENAME, data_only=True)
sheet_script = workbook["script"]
sheet_words = workbook["words"]
sheet_variables = workbook["variables"]
print("Читаем переменные..")
# Читаем переменные
variables_dict = {}
for row in sheet_variables.iter_rows(values_only=True):
    var_name, *var_values = row
    var_values = [var for var in var_values if var]
    variables_dict[var_name] = var_values
print("Чтение и обработка данных из листа script..")
# Чтение и обработка данных из листа script
for row in sheet_script.iter_rows(min_row=1, max_col=2):
    if row[0]:
        column_a = row[0].value.lower().strip() if row[0].value else ""
        selected_columns = [column_index_from_string(col) for col in column_a]
        result_array = []
        for col in selected_columns:
            multiplier = sheet_words.cell(row=3, column=col).value
            if not isinstance(multiplier, int):
                print(f"В {col} в третьей строке на листе words не цифра, пропускаем столбец")
                continue
            for cell in sheet_words.iter_rows(min_row=4, min_col=col, max_col=col, max_row=sheet_words.max_row, values_only=True):
                word = cell[0]
                if word:
                    # Умножаем слово на значение из второй строки
                    words_multiplied = [word] * abs(multiplier)
                        
                    # Обработка дефисов
                    for word_variant in words_multiplied:
                        result_array.append(word_variant)
            
        # Замена переменных
        processed_array = replace_words(variables_dict, result_array)

        # Формируем массив в формате JSON
        json_array = [f'"{word}"' for phrase  in processed_array for word in phrase.split()]
        json_array_2 = []
        for word in json_array:
            if "-" in word:
                json_array_2.append(word.replace("-", ""))
                json_array_2.append(word.replace("-", " "))
            else:
                json_array_2.append(word)
        json_result = "[" + ",".join(json_array_2) + "]"
        # Записываем результат в колонку B
        row[1].value = json_result

print("Сохраняем файл..")
# Сохраняем изменения
try:
    workbook.save(FILENAME)
    print(f"Файл {FILENAME} успешно сохранён..")
except PermissionError:
    print(f"Файл {FILENAME} не был сохранён. Попробуйте закрыть файл эксель перед запуском скрипта..")