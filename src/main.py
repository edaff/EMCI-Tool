# Eric Daff
# Created: 12/9/2019
# For Use By: EMCI Inc

from api_setup import gmail_api_setup
from datetime import datetime
import base64

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_query():
    return ("from:%s before:%s after:%s" %('celliott@emciwireless.com', int(datetime.now().timestamp()), int(datetime.today().replace(day=1).timestamp())))

def main():
    gmail_api = gmail_api_setup()

    query = get_query()

    response = gmail_api.users().messages().list(userId='me', q=query, maxResults=10).execute()

    for message_meta in response['messages']:
        print(message_meta['id'])
        message = gmail_api.users().messages().get(userId='me', id=response['messages'][0]['id'], format='raw').execute()
        msg_str = str(base64.urlsafe_b64decode(message['raw'].encode('ASCII')))
        print(msg_str)

if __name__ == '__main__':
    main()