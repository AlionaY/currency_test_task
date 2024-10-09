from collections import defaultdict
from datetime import date, datetime

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

credentials = 'creds.json'
sheet_id = '1CPvE_8clKDX2dxQgLv9LkOmcQYtp_XHRVeTeFqg9uj4'

worksheet_name = 'Currency test project'
default_cell_range_insert = 'A2'

dateformat = '%Y%m%d'
default_dateformat = '%Y-%m-%d'


def main():
    service_sheets = build_service_sheet()
    insert_row_index = get_last_filled_row_index(service_sheets) + 1
    cell_range = 'A' + str(insert_row_index)

    currency_response = get_currency_response_data()
    currency_info = merge_currency_data(currency_response)

    update_sheet_values(
        service_sheets,
        get_value_response_body(currency_info),
        cell_range=cell_range
    )


def get_default_date():
    return date.today()


def get_formatted_default_date(default_date=get_default_date()):
    """
    Change format of the default date to 'YYYYMMDD' for a GET request
    :param default_date: today (date object)
    :return: string
    """
    return str(default_date.strftime(default_dateformat))


def convert_string_to_date(currency_date_string):
    """
    Format date in string format to date object
    :param currency_date_string: date in string format to convert
    :return: date object
    """
    currency_date = datetime.strptime(currency_date_string, default_dateformat).date()
    return currency_date.strftime(dateformat)


def get_bank_url(
        update_from=get_formatted_default_date(),
        update_to=get_formatted_default_date()
):
    bank_url = 'https://bank.gov.ua/NBU_Exchange/exchange_site?start=' + update_from + \
               '&end=' + update_to + \
               '&sort=exchangedate&order=desc&json'
    return bank_url


def get_currency_response_data(update_from=get_formatted_default_date(),
                               update_to=get_formatted_default_date()):
    """
    Make a GET request and return a JSON file with currency data
    :param update_from: start date
    :param update_to: end date
    :return: JSON file
    """

    update_from_date = convert_string_to_date(update_from)
    update_to_date = convert_string_to_date(update_to)

    response = requests.get(get_bank_url(update_from_date, update_to_date))
    response_json = response.json()
    print('response json', response_json)
    return response_json


def merge_currency_data(response_json):
    """
    Merge data from a JSON to a one dict to write all values to the Google sheet
    :param response_json: a JSON from a request
    :return: dictionary
    """
    currency_info = defaultdict(list)

    for item in response_json:
        for key, value in item.items():
            currency_info[key].append(value)
    print('dict', currency_info)

    return currency_info


def get_value_response_body(currency_item: dict, dimension='COLUMNS'):
    values = tuple(currency_item.values())
    print('values:', values)
    value_response_body = {
        'majorDimension': dimension,
        'values': values
    }
    return value_response_body


def update_sheet_values(
        service_sheets,
        body,
        spreadsheet_id=sheet_id,
        value_input_option='USER_ENTERED',
        cell_range=default_cell_range_insert
):
    print('values to update', body)
    service_sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption=value_input_option,
        range=cell_range,
        body=body
    ).execute()


def build_service_sheet(filename=credentials):
    creds = service_account.Credentials.from_service_account_file(
        filename=filename
    )
    service_sheets = build('sheets', 'v4', credentials=creds)
    return service_sheets


def get_last_filled_row_index(service, spreadsheet_id=sheet_id):
    """
    Make GET response and find index of last filled row
    :param service:
    :param spreadsheet_id:
    :return: int
    """

    response = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Currency!A:A'
    ).execute()

    values = response.get('values', [])
    last_filled_row_index = len(values)
    print('filled row response', last_filled_row_index)
    return last_filled_row_index


if __name__ == '__main__':
    main()
