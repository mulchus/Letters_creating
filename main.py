from docxtpl import DocxTemplate
import os
import argparse
from environs import Env
import pandas
import string
from pathlib import Path


def get_addressees_from_file(addressees_file):
    addressees = pandas.read_excel(addressees_file, na_values='-', keep_default_na=False).to_dict('records')
    return addressees


def save_award(addressee, what_is_the_letter_about):
    # подставляем контекст в шаблон
    surname, name, patronymic = addressee['surname_name_patronymic'].split()
    surname_initials = f'{surname} {name[0]}. {patronymic[0]}.'
    initials_surname = f'{name[0]}. {patronymic[0]}. {surname}'
    addressee['surname_initials'] = surname_initials
    addressee['initials_surname'] = initials_surname
    addressee['name_patronymic'] = f'{name} {patronymic}'
    print(addressee)
    print(what_is_the_letter_about)
    doc = DocxTemplate("letter_tpl.docx")
    doc.render(addressee)

    # удаление знаков препинания из названия организации - исключить удаление "-"!!!
    organization = addressee['organization'].translate(str.maketrans('', '', string.punctuation))

    # !! сделать создание папок в самом начале - множество организаций из загруженного эксель - для ускорения
    file_path = Path.cwd() / 'Письма'  # / organization
    Path(file_path).mkdir(parents=True, exist_ok=True)
    print(f"{surname_initials} {what_is_the_letter_about}")
    doc.save(f"{file_path}/{surname_initials} {what_is_the_letter_about}.docx")



def main():
    docs_parser = argparse.ArgumentParser(description='Скрипт создания именных проектов писем')
    docs_parser.add_argument(
        'path',
        nargs='?',
        default=os.path.join(os.getcwd(), 'addressees.xlsx'),
        help='файлом адресатов *.xlsx (по умолчанию - ПУТЬ_К_ПАПКЕ_СО_СКРИПТОМ/addressees.xlsx)'
    )
    environs = Env()
    environs.read_env("setup.txt", recurse=False)
    docs = {}
    path = docs_parser.parse_args().path
    try:
        addressees = get_addressees_from_file(path)
    except (FileNotFoundError, ValueError) as error:
        print(f'Неверно указан путь к файлу.\nОшибка: {error}')
        print(f'Поиск в файле setup.txt в корневой папке.\n')
        try:
            path = environs.str('PATH_TO_AWARDS_FILE')
            addressees = get_addressees_from_file(path)
        except (FileNotFoundError, ValueError) as error:
            exit(f'setup.txt в корневой папке не найден или в нем не указан путь к файлу награждений.\n'
                 f'Ошибка: {error}')

    what_is_the_letter_about = environs.str('WHAT_IS_THE_LETTER_ABOUT')
    print(f'Скрипт запущен с файлом данных {path}')

    print('Создаем проекты писем')
    # addressees = get_addressees_from_file('addressees.xlsx')

    # определяем словарь переменных контекста, которые определены в шаблоне документа DOCX
    for addressee in addressees:
        save_award(addressee, what_is_the_letter_about)


if __name__ == '__main__':
    main()
