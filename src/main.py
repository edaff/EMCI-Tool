from api_setup import create_gmail_service, create_sheets_service, create_google_drive_service
from datetime import datetime
import base64

# Constants
SPREADSHEET_NAME = 'Dev Sample Rolling AR Report'
OWED_SHEET_NAME = 'Owed'
PAID_SHEET_NAME = 'Paid'
QUERY_EMAIL_ADDRESS = 'celliott@emciwireless.com'

def get_email_query():
    return ("from:%s before:%s after:%s" %(QUERY_EMAIL_ADDRESS, int(datetime.now().replace(day=31, month=12, year=2019).timestamp()), int(datetime.today().replace(day=1, month=12, year=2019).timestamp())))

def get_spreadsheet_query():
    return "name = '{0}'".format(SPREADSHEET_NAME)

def get_email_meta(gmail_service, log_file, response):
    meta = gmail_service.users().messages().get(userId='me', id=response['messages'][0]['id'], format='metadata').execute()
    log_file.write('Parsing Email: ' + extract_subject(meta) + '\n---------------------------------\n')

    return meta

def get_email_raw_data(gmail_service, response):
    message = gmail_service.users().messages().get(userId='me', id=response['messages'][0]['id'], format='raw').execute()
    
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
        amount = fields[4]

        # Grab the Transaction ID
        transcation_id_start_index = entry.find('Commission List ID = ') + len('Commission List ID = ')
        transaction_id = entry[transcation_id_start_index:end_index]

        # Store to the dictionary
        transactions[transaction_id] = amount

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

def main():
    num_mismatches = 0
    log_file = open('..\\logs\\' + datetime.now().strftime('%m-%d-%Y') + '.txt', 'w', encoding='utf-8')

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

    # Query for emails
    response = gmail_service.users().messages().list(userId='me', q=get_email_query(), maxResults=100).execute()

    # Throw an error and exit if no emails are found
    if(response['resultSizeEstimate'] == 0):
        log_file.write('ERROR: No emails found. Terminating program...')
        log_file.close()
        return

    for message_meta in response['messages']:
        num_mismatches = 0

        # Grab the message meta
        meta = get_email_meta(gmail_service, log_file, response)

        # Get the raw message string
        msg_str = get_email_raw_data(gmail_service, response)

        # Pull out the needed data
        transactions = get_transactions(msg_str)

        trace_number = get_trace_number(msg_str)

        transaction_date = get_transaction_date(msg_str)

        # Do spreadsheet processing


        if num_mismatches == 0:
            log_file.write('No Mismatches Found.\n---------------------------------\n\n')

    log_file.close()

if __name__ == '__main__':
    main()