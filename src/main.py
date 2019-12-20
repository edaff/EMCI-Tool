# Eric Daff
# Created: 12/9/2019
# For Use By: EMCI Inc

from api_setup import gmail_api_setup

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    gmail_api = gmail_api_setup()

    results = gmail_api.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])

if __name__ == '__main__':
    main()