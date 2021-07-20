# importing the required libraries
from datetime import datetime

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

COLUMNS = {
    'RECEIPT_NO': ('B1:F1', 'FACTURA NR. {RECEIPT_NO} Seria B'),

    'NAME': ('D2:F4', 'CUMPĂRĂTOR: {NAME}'),
    'CUI': ('D5:F5', 'CUI: {CUI}'),
    'ADDRESS': ('D6:F6', 'Adresă: {ADDRESS}'),
    'ACCOUNT': ('D7:F7', 'Cont: {ACCOUNT}'),
    'BANK': ('D8:F8', 'Banca: {BANK}'),

    'SERVICE_1_PRICE': ('C13', '{SERVICE_1_PRICE} RON'),
    'SERVICE_2_PRICE': ('C14', '{SERVICE_2_PRICE} RON'),
    'SERVICE_3_PRICE': ('C15', '{SERVICE_3_PRICE} RON'),

    'SERVICE_1_TOTAL': ('E13:F13', '{SERVICE_1_PRICE} RON'),
    'SERVICE_2_TOTAL': ('E14:F14', '{SERVICE_2_PRICE} RON'),
    'SERVICE_3_TOTAL': ('E15:F15', '{SERVICE_3_PRICE} RON'),

    'SUBTOTAL': ('E24:F24', '{SUBTOTAL}'),
    'TOTAL': ('D25:F25', 'Total: {TOTAL} RON'),
    'CONTRACT_INFO': ('A24:C25', 'Conform contractului nr. {CONTRACT_NO} din {CONTRACT_DATE}'),
}


def create_client():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('condica-320419-703d23b14fc2.json', scope)

    gc = gspread.authorize(creds)
    titles_list = []
    for spreadsheet in gc.openall():
        titles_list.append(spreadsheet.title)

    print(titles_list)

    return gc

def get_invoices_sheet(client, sheet_name='CONDICA'):
    sheet = client.open(sheet_name)
    return sheet.get_worksheet(0)

def get_invoice_data(client, sheet_name='CONDICA'):
    worksheet = get_invoices_sheet(client, sheet_name)
    records_data = worksheet.get_all_records()

    return pd.DataFrame.from_dict(records_data)


def create_invoice(client, data, index, tpl_name='tpl-factura'):
    date_now = datetime.now().strftime('%d.%m.%Y')
    invoice_name = 'FACTURA_{}_{}'.format(data['ID'], date_now)

    print('CREATED {}'.format(invoice_name))

    tpl_sheet = client.open(tpl_name)
    client.copy(tpl_sheet.id, title=invoice_name, copy_permissions=True)

    worksheet = client.open(invoice_name)

    data['SUBTOTAL'] = float(data['SERVICE_1_PRICE']) + float(data['SERVICE_2_PRICE']) + float(data['SERVICE_3_PRICE'])
    data['TOTAL'] = data['SUBTOTAL']

    sheet_instance = worksheet.get_worksheet(0)

    for key, item in COLUMNS.items():
        cell, format_str = item

        try:
            string_formatted = format_str.format(**data)
            sheet_instance.update(cell, string_formatted)
        except Exception as e:
            print(e)

    print('UPDATED {}'.format(invoice_name))
    worksheet.share(data['EMAIL'], perm_type='user', role='reader', notify=True)

    print('SHARED {}'.format(invoice_name))
    invoices_sheet = get_invoices_sheet(client)
    invoices_sheet.update_cell(index + 2, 14, invoice_name)
    invoices_sheet.update_cell(index + 2, 15, 'DA')

    print('SAVED {}'.format(invoice_name))
    return worksheet


client = create_client()

data = get_invoice_data(client)

for index, row in data.iterrows():
    if row.get('TRIMIS') != 'DA':
        print ('SEND ID: {} INVOICE'.format(row['ID']))
        create_invoice(client, row.to_dict(), index)
