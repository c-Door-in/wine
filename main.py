import os
import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import argparse
import pandas
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape


def set_year_word(age):
    if age[-2] != '1':
        if age[-1] == '1':
            return 'год'
        elif age[-1] > 1 and age[-1] < 5:
            return 'года'
    return 'лет'


def get_wine_cards_from_excel(file):
    wine_cards_from_excel = pandas.read_excel(
        io=file,
        na_values=None,
        keep_default_na=False,
    )
    return wine_cards_from_excel.to_dict(orient='records')


def sort_by_categories(source_table, column_name):
    categories = collections.defaultdict(list)
    for card in source_table:
        categories[card[column_name]].append(card)
    return categories


def main():
    load_dotenv()
    source_folder_path = os.getenv('SOURCE_FOLDER_PATH')
    xls_file_name = os.getenv('XLS_FILE_NAME')

    parser = argparse.ArgumentParser(
        description='Create index.html from template using data from wine excel table'
    )
    parser.add_argument('-src', '--source_file', default=xls_file_name)
    args = parser.parse_args()
    source_file = args.source_file

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    age = str(datetime.datetime.now().year - 1920)
    wine_cards = get_wine_cards_from_excel(f'{source_folder_path}{source_file}')

    rendered_page = template.render(
        years_old=age,
        year_word=set_year_word(age),
        sorted_wine_cards=sort_by_categories(wine_cards, column_name='Категория'),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
