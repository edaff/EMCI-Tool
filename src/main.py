# Eric Daff
# Created: 12/9/2019
# For Use By: EMCI Inc

from api_setup import create_gmail_service, create_sheets_service
from datetime import datetime
import base64

def get_query():
    return ("from:%s before:%s after:%s" %('celliott@emciwireless.com', int(datetime.now().timestamp()), int(datetime.today().replace(day=1).timestamp())))

def main():
    gmail_service = create_gmail_service()

    query = get_query()

    response = gmail_service.users().messages().list(userId='me', q=query, maxResults=10).execute()

    for message_meta in response['messages']:
        print(message_meta['id'])
        message = gmail_service.users().messages().get(userId='me', id=response['messages'][0]['id'], format='raw').execute()
        msg_str = str(base64.urlsafe_b64decode(message['raw'].encode('ASCII')))
        print(msg_str)

if __name__ == '__main__':
    main()