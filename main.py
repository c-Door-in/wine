from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import collections
import pandas
import datetime


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

def years_from_start():
    start = datetime.datetime(year=1920, month=1, day=1)
    now = datetime.datetime.now()
    delta = now.year - start.year
    return delta

def set_year_word():
    years_amount = str(years_from_start())
    if years_amount[-2] != '1':
        if years_amount[-1] == '1':
            return 'год'
        elif years_amount[-1] > 1 and years_amount[-1] < 5:
            return 'года'
    return 'лет'

def get_wine_cards_from_excel(file):
    wine_from_excel = pandas.read_excel(
        io=file,
        na_values=None,
        keep_default_na=False,
    )
    return wine_from_excel.to_dict(orient='records')

wine_dict = get_wine_cards_from_excel('input/wine3.xlsx')

def sort_by_categories(source_dict, column_name):
    sort_by_categories = collections.defaultdict(list)
    for card in source_dict:
        sort_by_categories[card[column_name]].append(card)
    return sort_by_categories


rendered_page = template.render(
    years_old=years_from_start(),
    year_word=set_year_word(),
    sorted_wine_dict=sort_by_categories(wine_dict, column_name='Категория'),
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
