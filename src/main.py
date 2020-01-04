# Eric Daff
# Created: 12/9/2019
# For Use By: EMCI Inc

from api_setup import create_gmail_service, create_sheets_service
from datetime import datetime
import base64

def get_query():
    return ("from:%s before:%s after:%s" %('celliott@emciwireless.com', int(datetime.now().replace(day=31, month=12, year=2019).timestamp()), int(datetime.today().replace(day=1, month=12, year=2019).timestamp())))

def get_transactions(msg_str):
    transactions = {}

    # Find the start of the transactions
    start_index = msg_str.find('Motorola Solutions Inc\\r\\n') + len('Motorola Solutions Inc\\r\\n')
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

def extract_subject(meta):
    headers = meta['payload']['headers']

    for header in headers:
        if(header['name'] == 'Subject'):
            return header['value']

def main():
    num_mismatches = 0
    log_file = open('..\\logs\\' + datetime.now().strftime('%m-%d-%Y') + '.txt', 'w', encoding='utf-8')

    gmail_service = create_gmail_service()
    query = get_query()
    response = gmail_service.users().messages().list(userId='me', q=query, maxResults=10).execute()

    for message_meta in response['messages']:
        num_mismatches = 0
        meta = gmail_service.users().messages().get(userId='me', id=response['messages'][0]['id'], format='metadata').execute()
        log_file.write('Parsing Email: ' + extract_subject(meta) + '\n---------------------------------\n')

        message = gmail_service.users().messages().get(userId='me', id=response['messages'][0]['id'], format='raw').execute()
        msg_str = str(base64.urlsafe_b64decode(message['raw'].encode('ASCII')))
        transactions = get_transactions(msg_str)
        print(transactions)

        if num_mismatches == 0:
            log_file.write('No Mismatches Found.\n---------------------------------\n\n')

    log_file.close()

if __name__ == '__main__':
    main()