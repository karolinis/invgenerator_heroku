import pandas as pd
import requests
import glob
import pandas._libs.tslibs.base
import openpyxl
import streamlit as st
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import warnings
import base64

def sender(customer_name, payload):
    header = {'X-AppSecretToken' : 'rIMJ7rEcyp9h88IIBUHXMi9T2JeVYWW9HQtpV1P5a7c1',
              'X-AgreementGrantToken' : 'BS6I6xBFgUykx7IjKBCaFLhoL7mfkM5MeFM7w69hfDY1',
              'Content-Type' : 'application/json'}

    url = 'https://restapi.e-conomic.com/invoices/drafts'

    # Calling API POST method to create invoice
    r = requests.post(url, headers=header, json=payload)
    print(r.status_code, r.reason, f'Invoice generated for {customer_name}')

def create_json(customer_no, customer_name, heading, text_line1,date):
    payload = {
        'date': date,
        'customer': {'customerNumber': customer_no},
        'layout': {'layoutNumber': 19},
        'currency': 'DKK',
        'paymentTerms': {'paymentTermsNumber': 1},
        'recipient': {
            'name': customer_name,
            'vatZone': {'name': 'Domestic', 'vatZoneNumber': 1}
        },
        'notes': {
            'heading': heading,
            'textLine1': text_line1,

        },
        'lines': []
    }
    return payload

def create_line(payload, line_no, period = None):
    payload['lines'] = [
                        {
                        "lineNumber": line_no,
                        'description': period,
                        }]
    return payload, line_no + 1

def append_line(payload, line_no, product_no = None, product_name = None, quantity = None, price = None, period = None):
    payload['lines'].append({
                            "lineNumber": line_no,
                            'description': product_name,
                            'product': {'productNumber': product_no},
                            'quantity': round(float(quantity),0),
                            'unitNetPrice': round(float(price),2),
                            })
    return payload, line_no


if __name__ == '__main__':
    user_inpt = st.file_uploader('Choose file to import')
    if not user_inpt:
        st.stop()
        st.success('Thank you')

    date_input = st.text_input('Enter date for invoices (yyyy-mm-dd): ')
    if not date_input:
        st.stop()
        st.success('Thank you')

    warnings.filterwarnings('ignore')

    dataframe = pd.read_excel(user_inpt)
    # create a dictionary invoices which will store customers and invoice period as keys
    # inside we will store all needed info to upload to e-conomic
    invoices = {}
    for i, value in dataframe.iterrows():
        if f'{value["Customer no"]}' not in invoices:
            invoices[f'{value["Customer no"]}'] = [value.values.tolist()]
        else:
            invoices[f'{value["Customer no"]}'].append(value.values.tolist())

    # customer_no, customer_name, line_no, product_no, product_name, quantity, heading, text1
    st.write('Sent to: ')
    for invoice in invoices:
        # inside of this loop we will create new variables to make it easier to read
        # when period changes, we create a line for the period and then invoice lines
        line_no = 1
        lines_payload = {}  # keys are periods, this helps sort the lines
        for i in invoices[invoice]:
            customer_no, customer_name, headline, text_line1, period, product_no, product_name, quantity, price, total = \
                i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]
            if str(text_line1) == 'nan' and str(period) == 'nan':
                text_line1 = ' '
                period = ' '
            payload = create_json(customer_no, customer_name, headline, text_line1,date_input)
            if str(period) == 'nan':
                new_payload, new_line_no = append_line(payload, line_no, product_no, product_name, quantity,
                                                       price)
                lines_payload[period] = new_payload['lines']

            else:
                if period not in lines_payload:
                    payload, new_line_no = create_line(payload, line_no, period)
                    new_payload, new_line_no = append_line(payload, new_line_no, product_no, product_name, quantity,
                                                           price,
                                                           period)
                    lines_payload[period] = new_payload['lines']
                else:
                    new_payload, new_line_no = append_line(new_payload, line_no, product_no, product_name, quantity,
                                                           price)
                    lines_payload[period] = new_payload['lines']
            line_no = new_line_no + 1

        final_payload = {}

        for i in lines_payload:
            if 'lines' not in final_payload:
                final_payload['lines'] = lines_payload[i]
            else:
                final_payload['lines'] = final_payload['lines'] + lines_payload[i]

        new_payload['lines'] = final_payload['lines']
        sender(customer_name, new_payload)
        st.write(customer_name)
    st.write('Program has successfully finished')