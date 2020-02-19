from api_setup import create_gmail_service, create_sheets_service, create_google_drive_service
from datetime import datetime, timedelta
import base64
from enum import Enum
import re
import decimal

# Constants
SPREADSHEET_NAME = 'Rolling AR Report'
OWED_SHEET_NAME = 'Owed'
PAID_SHEET_NAME = 'Paid'
QUERY_EMAIL_ADDRESS = 'comtech@emciwireless.com'
BEFORE_DATE = int(datetime.now().timestamp())
AFTER_DATE = int((datetime.today() - timedelta(days=5)).timestamp())

class COLUMN_NAMES(Enum):
    TRANSACTION_ID = 'Transaction ID'
    OWED_TO_COM_TECH = 'Owed to Com-Tech'
    PAID_TO_DATE = 'Paid To Date'
    DATE_PAID = 'Date Paid'
    CHECK_NUMBER = 'Check number'

def get_email_query():
    return ("from:%s before:%s after:%s" %(QUERY_EMAIL_ADDRESS, BEFORE_DATE, AFTER_DATE))

def get_spreadsheet_query():
    return "name = '{0}'".format(SPREADSHEET_NAME)

def get_email_meta(gmail_service, log_file, message_meta):
    meta = gmail_service.users().messages().get(userId='me', id=message_meta['id'], format='metadata').execute()
    email_subject = extract_subject(meta)
    log_file.write('Processing Email: ' + email_subject + '\n---------------------------------\n')
    print('Processing Email: ' + email_subject + '\n')

    return meta

def get_email_raw_data(gmail_service, message_meta):
    message = gmail_service.users().messages().get(userId='me', id=message_meta['id'], format='raw').execute()
    
    return str(base64.urlsafe_b64decode(message['raw'].encode('ASCII')))

# Fetch all transactions from the email
def get_transactions(msg_str):
    transactions = {}

    # Find the start of the transactions
    start_index = msg_str.find('\\r\\n', msg_str.find('ENTITY:')) + len('\\r\\n')
    end_index = msg_str.find('\\\\', start_index, -1)
    entry = msg_str[start_index:end_index]

    # Iterate over the email and grab each Transaction ID & Amount pair
    while end_index != -1:
        # Grab the Transation Amount
        fields_end_index = entry.find('\\r\\n')
        fields = entry[0:fields_end_index].split()

        if(len(fields) == 0):
            break

        raw_amount = fields[4]
        formatted_amount = '$' + f'{decimal.Decimal(raw_amount):,}'

        # Grab the Transaction ID
        if(entry.find('List ID = ') != -1):
            transcation_id_start_index = entry.find('List ID = ') + len('List ID = ')
        else:
            transcation_id_start_index = entry.find('List ID: ') + len('List ID: ')

        transaction_id = entry[transcation_id_start_index:end_index]

        # Store to the dictionary
        transactions[transaction_id] = formatted_amount

        # Move to the next transaction
        start_index = msg_str.find('\\\\\\r\\n\\r\\n', start_index) + len('\\\\\\r\\n\\r\\n')
        end_index = msg_str.find('\\\\', start_index)
        entry = msg_str[start_index:end_index]

    return transactions

# Scrape the trace number from the email
def get_trace_number(msg_str):
    trace_number_start_index = msg_str.find('PAYOR TRANSACTION TRACE #:') + len('PAYOR TRANSACTION TRACE #:')
    trace_number_end_index = msg_str.find('\\r\\n', trace_number_start_index)
    
    return msg_str[trace_number_start_index:trace_number_end_index].strip()

# Scrape the transaction date from the email
def get_transaction_date(msg_str):
    transaction_date_start_index = msg_str.find('USD SCHEDULED TO SETTLE OR PAY ON') + len('USD SCHEDULED TO SETTLE OR PAY ON')
    transaction_date_end_index = msg_str.find('**', transaction_date_start_index)

    date_string = msg_str[transaction_date_start_index:transaction_date_end_index].strip().strip('\\r\\n')
    actual_date = datetime.strptime(date_string, '%y/%m/%d')
    
    return datetime.strftime(actual_date, '%m/%d/%y')

# Extract the subject from the email for logs
def extract_subject(meta):
    headers = meta['payload']['headers']

    for header in headers:
        if(header['name'] == 'Subject'):
            return header['value']

def find_column(sheet, column_name):
    return sheet.get('values', [])[0].index(column_name)

def get_sheet_id(sheet_meta, sheet_name):
    for sheet in sheet_meta['sheets']:
        if(sheet['properties']['title'] == sheet_name):
            return sheet['properties']['sheetId']

def find_transaction_row(transaction_id_column, sheet_values, transaction):
    for idx, value in enumerate(sheet_values):
        if(len(value) == 0):
            continue
        if(value[transaction_id_column] == transaction):
            return idx

def copy_row_request(transaction_row, owed_sheet_values, paid_sheet_values, spreadsheet_id, sheets_service):
    insert_request_body = {
        "values": [
            owed_sheet_values[transaction_row][0:15]
        ]
    }
    request_range = "'{0}'!A{1}:P{2}".format(PAID_SHEET_NAME, len(paid_sheet_values) + 1, len(paid_sheet_values) + 1)

    return sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=request_range,
        valueInputOption='USER_ENTERED',
        includeValuesInResponse=False,
        body=insert_request_body).execute()

def delete_row_request(owed_sheet_id, transaction_row, spreadsheet_id, sheets_service):
    delete_request_body = {
        "requests": [
            {
            "deleteDimension": {
                "range": {
                "sheetId": owed_sheet_id,
                "dimension": "ROWS",
                "startIndex": transaction_row,
                "endIndex": transaction_row + 1
                }
            }
            }
        ],
        }

    return sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=delete_request_body).execute()

def strip_non_numbers(val):
    return re.sub("[^0-9]", "", val)

def main():
    num_mismatches = 0
    log_file_name = '..\\logs\\' + datetime.now().strftime('%m-%d-%Y') + '.txt'
    log_file = open(log_file_name, 'w', encoding='utf-8')

    # ********************************
    #           API Setup
    # ********************************

    # Create services for APIs
    gmail_service = create_gmail_service()
    sheets_service = create_sheets_service()
    drive_service = create_google_drive_service()

    # Fetch the correct spreadsheet
    response = drive_service.files().list(q=get_spreadsheet_query()).execute()

    # Throw an error if the spreadsheet can't be found
    if(response['files'] == []):
        log_file.write('ERROR: Unable to find spreadsheet. Terminating program...')
        log_file.close()
        return

    # Query for the spreadsheet using the found spreadsheet id
    log_file.write('Found spreadsheet: ' + response['files'][0]['name'] + '\n\n')
    spreadsheet_id = response['files'][0]['id']

    sheet_meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=False).execute()
    owed_sheet_id = get_sheet_id(sheet_meta, OWED_SHEET_NAME)
    paid_sheet_id = get_sheet_id(sheet_meta, PAID_SHEET_NAME)

    # Query for emails
    response = gmail_service.users().messages().list(userId='me', q=get_email_query(), maxResults=100).execute()

    # Throw an error and exit if no emails are found
    if(response['resultSizeEstimate'] == 0):
        log_file.write('ERROR: No emails found. Terminating program...')
        log_file.close()
        return

    # ********************************
    #      Main Program Loop
    # ********************************

    for message_meta in response['messages']:
        num_mismatches = 0
        num_missing_rows = 0

        # ********************************
        #         Email Scraping
        # ********************************

        # Grab the message meta
        meta = get_email_meta(gmail_service, log_file, message_meta)

        if("EDI ADVICE" not in extract_subject(meta)):
            continue

        # Get the raw message string
        msg_str = get_email_raw_data(gmail_service, message_meta)

        # Pull out the needed data
        transactions = get_transactions(msg_str)
        trace_number = get_trace_number(msg_str)
        transaction_date = get_transaction_date(msg_str)

        # ********************************
        #      Spreadsheet Processing
        # ********************************

        # Iterate over all of the transactions
        for transaction in transactions:
            # Fetch the two needed sheets and their data
            owed_sheet = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=OWED_SHEET_NAME).execute()
            paid_sheet = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=PAID_SHEET_NAME).execute()
            owed_sheet_values = owed_sheet.get('values', [])
            paid_sheet_values = paid_sheet.get('values', [])

            # Grab indexes for all of the required columns
            transaction_id_column = find_column(owed_sheet, COLUMN_NAMES.TRANSACTION_ID.value)
            owed_to_com_tech_column = find_column(owed_sheet, COLUMN_NAMES.OWED_TO_COM_TECH.value)
            paid_to_date_column = find_column(owed_sheet, COLUMN_NAMES.PAID_TO_DATE.value)
            date_paid_column = find_column(owed_sheet, COLUMN_NAMES.DATE_PAID.value)
            check_number_column = find_column(owed_sheet, COLUMN_NAMES.CHECK_NUMBER.value)

            # Find the row with the corresponding transaction id
            transaction_row = find_transaction_row(transaction_id_column, owed_sheet_values, transaction)

            # Skip this transaction if the corresponding row can't be found in the Owed sheet
            if(transaction_row == None):
                if(find_transaction_row(transaction_id_column, paid_sheet_values, transaction) == None):
                    log_file.write('Missing Transaction Found: Transaction ID {0} NOT FOUND in owed or paid sheet ... Skipping\n'.format(transaction))
                    num_missing_rows += 1
                continue

            # If the transaction ID already exists in the paid sheet, this guy has already been processed
            if(find_transaction_row(transaction_id_column, paid_sheet_values, transaction) != None):
                continue

            # Skip this transaction if the Owed to Com-Tech value does not match the transaction value
            if(strip_non_numbers(owed_sheet_values[transaction_row][owed_to_com_tech_column]) != strip_non_numbers(transactions[transaction])):
                log_file.write("Mismatch Found: Transaction ID {0} 'Owed to Com-Tech' value of '{1}' does not match the transaction's value of '{2}' ... Skipping\n".format(
                    transaction,
                    owed_sheet_values[transaction_row][owed_to_com_tech_column],
                    transactions[transaction]))
                num_mismatches += 1
                continue

            paid_sheet_size = len(paid_sheet_values) + 1

            # Format the new spreadsheet row to be inserted into the 'Paid' sheet
            owed_sheet_values[transaction_row][owed_to_com_tech_column] = "=I{0}-K{1}".format(paid_sheet_size, paid_sheet_size)
            owed_sheet_values[transaction_row][paid_to_date_column] = transactions[transaction]
            owed_sheet_values[transaction_row][date_paid_column] = transaction_date
            owed_sheet_values[transaction_row][check_number_column] = trace_number

            # Copy the found row to the paid sheet
            copy_row_request(transaction_row, owed_sheet_values, paid_sheet_values, spreadsheet_id, sheets_service)

            # Delete the old row from the Owed sheet
            delete_row_request(owed_sheet_id, transaction_row, spreadsheet_id, sheets_service)
            

        log_file.write('---------------------------------\n{0} Mismatches Found.\n---------------------------------\n'.format(num_mismatches))
        log_file.write('{0} Missing Transactions Found.\n---------------------------------\n\n'.format(num_missing_rows))


    print('Processing Complete! Logs saved to: {0}'.format(log_file_name))
    log_file.close()

if __name__ == '__main__':
    main()