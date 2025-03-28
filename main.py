import requests

from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_table_elements(table_currencies):
    """
    table_elements_dict -- Список с название-курс идущими подряд:
                                    ['Австралийский доллар', '52,9248', 'Азербайджанский манат', '49,3145, ...']
    table_elements -- Список table_elements_dict разбитый на список с попарными вложенными списками название-курс:
                                    [['Австралийский доллар', '52,9248'], ['Азербайджанский манат', '49,3145'], ...]


    :param: table_currencies
    :return: table_elements 
    """
    table_elements_dict =[]
    for item in table_currencies:
        table_element = item.find_all("td")[3:]
        for el in table_element:
            table_elements_dict.append(el.text)

    table_elements = [table_elements_dict[i:i+2] for i in range(0, len(table_elements_dict), 2)]

    return table_elements


def save_to_exel(title, currency_headers, currencies):
    wb = Workbook()
    ws = wb.active

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 9

    ws['A1'] = title
    ws.append(currency_headers)
    for currencie in currencies:
        ws.append(currencie)



    wb.save("Currencies.xlsx")


def main():
    """
    title -- Текст описания с датой
    currency_headers -- Список с заголовками ['Валюта', 'Курс']
    currencies -- Список с элементами [[название, курс], [название, курс], ...]

    :return: None
    """

    url = "https://www.cbr.ru/currency_base/daily/"

    response = requests.get(url)
    soup = BeautifulSoup(response.text, "lxml")

    table = soup.find('table').find("tbody")
    table_currencies = (table.find_all('tr')[1:]) #Список со всеми валютами

    # Данные для переноса в Exel
    title = soup.find("h2", class_='h3').text
    currency_headers = table.find('tr').text.split()[5:]
    currencies = get_table_elements(table_currencies)

    save_to_exel(title, currency_headers, currencies)



if __name__ == '__main__':
    main()