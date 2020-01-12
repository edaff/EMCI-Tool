# EMCI-Tool
# Created by: Eric Daff
# For Use By: EMCI Wireless Inc

# To Run:
1. Install Python 3 (https://www.python.org/ftp/python/3.8.1/python-3.8.1.exe)
- Ensure you select the "Add Python to Path" option
2. Login to the Google Cloud Developer Manager (https://console.developers.google.com/apis/dashboard?folder=&organizationId=&project=emci-tool)
- Create a new cloud project
- Click 'Enable APIs and services'
- Search for and add 3 apps: Google Drive, Gmail and Google Sheets
- Go back to the main page for your project
- Click 'Credentials'
- Click 'Create Credentials' -> 'Oauth Client ID'
- Check the 'Other' box, enter a name, and hit create
- Hit the download button on the end of the bar for your new oath client id
- Save it to the 'auth' folder in the project directory and name it 'credentials.json'
3. Run 'EMCI-Tool.bat' inside of the bin folder
- Logs will be located in 'logs/{todays-date}.txt'


# Resources:
## Gmail
- https://developers.google.com/gmail/api/quickstart/python
- https://developers.google.com/gmail/api/guides/filtering

## Sheets
- https://developers.google.com/sheets/api/guides/concepts
- https://developers.google.com/sheets/api/samples/sheet
- https://developers.google.com/sheets/api/samples/sheet#determine_sheet_id_and_other_properties

## Drive
- https://developers.google.com/drive/api/v3/search-files