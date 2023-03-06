from docxtpl import DocxTemplate
# import argparse
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
        doc.save(f"{file_path}/{surname_initials} {organization} {award['award_type']} протокол.docx")


def main():

    task_id = int(input('1 - Создание Почетных граммот и Благодарностей\n '
                    '2 - Создание выписок из протоколов\n'
                    'Введите задачу: \n '))

    if task_id == 1:
        awards = get_awards_from_file('awards.xlsx')
    elif task_id == 2:
        awards = get_awards_from_file('awards_for_protocol.xlsx')
    else:
        exit('Что то не так')

    # определяем словарь переменных контекста, которые определены в шаблоне документа DOCX
    i = 1
    for award in awards:
        save_award(award, task_id)
        i += 1
        if i > 5:
            break

    # реализовать указание файла со списком награждаемых, по умолчанию - awards.xlsx
    # wine_parser = argparse.ArgumentParser(description='Сайт магазина авторского вина "Новое русское вино"')
    # wine_parser.add_argument(
    #     'path',
    #     nargs='?',
    #     default=os.path.join(os.getcwd(), 'wine.xlsx'),
    #     help='директория с файлом wine.xlsx (по умолчанию - ПУТЬ_К_ПАПКЕ_СО_СКРИПТОМ/wine.xlsx)'
    # )
    #
    # path = wine_parser.parse_args().path
    # try:
    #     wines = get_wines_from_file(path)
    # except (FileNotFoundError, ValueError) as error:
    #     print(f'Неверно указан путь к файлу.\nОшибка: {error}')
    #     environs = Env()
    #     try:
    #         environs.read_env("setup.txt", recurse=False)
    #         path = environs.str('PATH_TO_WINE_FILE')
    #         wines = get_wines_from_file(path)
    #     except (FileNotFoundError, ValueError) as error:
    #         print(f'setup.txt в корневой папке не найден или в нем не указан путь к файлу вин в PATH_TO_WINE_FILE.\n'
    #               f'Ошибка: {error}')
    #         exit()
    #
    # print(f'Сайт запущен с файлом базы данных {path}')


if __name__ == '__main__':
    main()
