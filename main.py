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


def group_cards_to_categories(cards, categories_column_name):
    groups_by_categories = collections.defaultdict(list)
    for card in cards:
        groups_by_categories[card[categories_column_name]].append(card)
    return groups_by_categories


def add_missing_args_to_parser(vars_in_dotenv):
    parser = argparse.ArgumentParser(
        description='Create index.html from template using data from wine excel table'
    )
    for var_name, var_value in vars_in_dotenv.items():
        if var_value:
            parser.add_argument(
                f'--{var_name}',
                default = var_value,
                help=f'Value in ".env" is "{var_value}". You can change this argument here.'
            )
        else:
            parser.add_argument(
                f'{var_name}',
                help = f'This argument is not in ".env". You can point it here.'
            )
    return parser.parse_args()


def get_vars_from_dotenv(*dotenv_vars_names):
    load_dotenv()
    return {var_name: os.getenv(var_name) for var_name in dotenv_vars_names}


def main():
    vars_from_dotenv = get_vars_from_dotenv(
        'SOURCE_FOLDERPATH',
        'SOURCE_FILENAME',
    )
    args = add_missing_args_to_parser(vars_from_dotenv)
    source_filepath = '{SOURCE_FOLDERPATH}/{SOURCE_FILENAME}'.format(**vars(args))

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    founded_year = 1920
    age = datetime.datetime.now().year - founded_year
    wine_cards = get_wine_cards_from_excel(source_filepath)

    rendered_page = template.render(
        years_old=age,
        year_word=set_year_word(age),
        wine_cards_groups=group_cards_to_categories(wine_cards, 'Категория'),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
