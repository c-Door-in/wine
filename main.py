import argparse
import collections
import datetime
import os
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape


def set_year_word(age):
    age = str(age)
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


def group_cards_to_categories(source_table, column_name):
    categories = collections.defaultdict(list)
    for card in source_table:
        categories[card[column_name]].append(card)
    return categories


def main():
    load_dotenv()
    source_folder_path = os.getenv('SOURCE_FOLDER_PATH')
    xls_filename = os.getenv('XLS_FILENAME')

    parser = argparse.ArgumentParser(
        description='Create index.html from template using data from wine excel table'
    )
    parser.add_argument('-src', '--source_filename', default=xls_filename)
    args = parser.parse_args()
    source_filename = args.source_filename

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    founded_year = 1920
    age = datetime.datetime.now().year - founded_year
    wine_cards = get_wine_cards_from_excel(f'{source_folder_path}{source_filename}')

    rendered_page = template.render(
        years_old=age,
        year_word=set_year_word(age),
        wine_cards_groups=group_cards_to_categories(wine_cards, column_name='Категория'),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
