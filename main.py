from docxtpl import DocxTemplate
import os
import argparse
from environs import Env
import pandas
import string
from pathlib import Path


def get_awards_from_file(awards_file):
    awards = pandas.read_excel(awards_file, na_values='-', keep_default_na=False).to_dict('records')
    return awards


def save_award(award, task_id):
    # подставляем контекст в шаблон
    if task_id == 1 and award['award_type'] == 'Почетная грамота':
        doc = DocxTemplate("diploma_tpl.docx")
    elif task_id == 1 and award['award_type'] == 'Благодарность':
        doc = DocxTemplate("gratitude_tpl.docx")
    elif task_id == 2:
        doc = DocxTemplate("protocol_tpl.docx")
    else:
        return

    doc.render(award)

    surname, name, patronymic = award['surname_name_patronymic'].split()
    surname_initials = f'{surname} {name[0]}.{patronymic[0]}.'

    # удаление знаков препинания из названия организации - исключить удаление "-"!!!
    organization = award['organization'].translate(str.maketrans('', '', string.punctuation))

    # !! сделать создание папок в самом начале - множество организаций из загруженного эксель - для ускорения
    file_path = Path.cwd() / 'Награды' / organization
    Path(file_path).mkdir(parents=True, exist_ok=True)

    print(f"{surname_initials} {organization} {award['award_type']}")
    if task_id == 1:
        doc.save(f"{file_path}/{surname_initials} {organization} {award['award_type']}.docx")
    elif task_id == 2:
        doc.save(f"{file_path}/{surname_initials} {organization} протокол.docx")


def main():
    award_parser = argparse.ArgumentParser(description='Скрипт создания проектов бланков награждений или '
                                                       'выписок из протоколов')
    award_parser.add_argument(
        'path',
        nargs='?',
        default=os.path.join(os.getcwd(), 'awards.xlsx'),
        help='директория с файлом awards.xlsx или awards_for_protocol.xlsx '
             '(по умолчанию - ПУТЬ_К_ПАПКЕ_СО_СКРИПТОМ/awards.xlsx)'
    )
    award_parser.add_argument(
        'task_id',
        nargs='?',
        default=1,
        help='Тип задачи: \n'
             '1 - Создание Почетных граммот и Благодарностей\n'
             '2 - Создание выписок из протоколов\n'
    )

    awards = {}
    path = award_parser.parse_args().path
    try:
        awards = get_awards_from_file(path)
    except (FileNotFoundError, ValueError) as error:
        print(f'Неверно указан путь к файлу.\nОшибка: {error}')
        print(f'Поиск в файле setup.txt в корневой папке.\n')
        environs = Env()
        try:
            environs.read_env("setup.txt", recurse=False)
            path = environs.str('PATH_TO_AWARDS_FILE')
            awards = get_awards_from_file(path)
        except (FileNotFoundError, ValueError) as error:
            exit(f'setup.txt в корневой папке не найден или в нем не указан путь к файлу награждений.\n'
                 f'Ошибка: {error}')

    print(f'Скрипт запущен с файлом данных {path}')

    task_id = int(award_parser.parse_args().task_id)
    if task_id not in (1, 2):
        task_id = int(input('1 - Создание Почетных граммот и Благодарностей\n'
                            '2 - Создание выписок из протоколов\n'
                            'Введите задачу: \n '))

    if task_id == 1:
        print('Создаем проекты грамот и благодарностей')
        awards = get_awards_from_file('awards.xlsx')
    elif task_id == 2:
        print('Создаем проекты выписок из протоколов')
        awards = get_awards_from_file('awards_for_protocol.xlsx')
    else:
        exit('Неверно передан код задачи')

    # определяем словарь переменных контекста, которые определены в шаблоне документа DOCX
    for award in awards:
        save_award(award, task_id)


if __name__ == '__main__':
    main()
