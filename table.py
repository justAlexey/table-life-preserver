
import pandas
import settings
import os

previous_cm_number = ""
path_to_archive = settings.path_to_archive

def get_file_list() -> list:
    """возвращает кортеж со списком файлов в папке Archive"""
    print("Пытаюсь найти папку с проектами...")
    try:
        archive = os.walk(path_to_archive)
        return list(archive)[0][2]
    except:
        print("Не могу найти папку \"Archive\"")
        input("Нажмите любую кнопку, что бы выйти...")
        os._exit(0)

def get_serial_numbers_by_file(file: str)->list:
    """Возвращает список серийных номеров из файла"""
    print(f"Извлекаю номера из {file}")
    temp = pandas.read_excel(f"{path_to_archive}{file}")
    temp = temp.iloc[6:25, [3]].squeeze().tolist()
    return ["L"+str(x) for x in temp]

def get_formated_list(cm: str)->str:
    """Возвращает очищенную строку со списком серийных номеров"""
    print(f"Формирую список серийных номеров {cm}")
    serial_numbers = get_serial_numbers_by_file(cm)
    for item in serial_numbers[::-1]:
        if item == 'Lnan':
            serial_numbers.remove('Lnan')
    return ", ".join(serial_numbers)

def format_numbers_from_cm(cm):
    """Возвращает сформированную строку с номером СМ и списком серийных номеров"""
    print(f"Формирую строку для письма {cm}")
    
    return f"{cm.replace('.xlsm','')}:\nS/N: {get_formated_list(cm)}\n\n"

def filter_cm_numbers()->list:
    """Фильтрует список СМ проектов от открытых файлов"""
    print("Фильтрую список проектов...")
    cm_projects = get_file_list()
    for item in cm_projects[::-1]:
        if item[0] == '~':
            cm_projects.remove(item)
    return cm_projects

def main()->None:
    print("Начинаю...")
    with open("Сгенерированное письмо на добавление серийных номеров.txt", "w") as file:
        file.write(f"Коллеги, прошу добавить серийные номера в заявке проекта: ")
        for cm in filter_cm_numbers():
            file.write(format_numbers_from_cm(cm))
    # print(filter_cm_numbers())
    print("Письмо готово. Смотри в файле \"Сгенерированное письмо на добавление серийных номеров.txt\"")
    input("Для закрытия окна нажмите \"Enter\"...")
if __name__ == "__main__":
    main()