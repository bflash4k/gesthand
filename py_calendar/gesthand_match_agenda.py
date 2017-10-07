''' -----------------------------------------------------------------------
Module python pour importer les dates des matchs dans un agenda google calendar
'''

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from googleapiclient.errors import HttpError

import datetime
import sys
import getopt
import csv
from datetime import datetime, timedelta
import codecs

all_matchs = []
team_list = []

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
#SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret_gcalendar.json'
APPLICATION_NAME = 'Gesthand Calendar helper'

# -------------------------------------------------------------
def get_credentials(flags):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'calendar-gesthand.json')

    store = Storage(credential_path)
    #pb ci dessous...
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

#==============================================================
def StoreMyEvents(service):
  for lines in all_matchs:
    ''' This is a dictionnary fill with values
    '''
    event = {}
    event['start']= {}
    event['end']= {}

    #fill the values
    event['summary'] = lines['club rec']+"/"+lines['club vis']
    event['location'] = lines['nom salle']+","+ lines['adresse salle']+","+ lines['CP']+","+ lines['Ville']
    event['description'] = "J"+lines['J']+" "+lines['competition']

    # Convert date format from dd/mm/yyyy to yyyy-mm-dd
    d = datetime.strptime(lines['le']+" "+lines['horaire'], '%d/%m/%Y %H:%M:%S')
    when = d.strftime('%Y-%m-%dT%H:%M:%S')
    event['start']['dateTime'] = d.strftime('%Y-%m-%dT%H:%M:%S')   #'2017-08-28T09:00:00-07:00'
    event['start']['timeZone'] = 'Europe/Paris'

    #it last 2 hours
    d +=  timedelta(hours=2)
    event['end']['dateTime'] = d.strftime('%Y-%m-%dT%H:%M:%S')
    event['end']['timeZone'] = 'Europe/Paris'

    #event['colorId'] = 'red'
    #unique ID !!!
    tag = lines['num poule']+lines['J']
    event['id'] = tag.lower()

    try:
      event['colorId'] = team_list.index(lines['poule'])

    except ValueError:
      team_list.append(lines['poule'])
    
    #start my colors from 4
    event['colorId'] = 4 + team_list.index(lines['poule'])
    print(event)

    try:
      event = service.events().insert(calendarId='primary', body=event).execute()

    except HttpError as err:

      if (err.resp.status == 409):  #already exist
        #update the event and retry
        print("Updated")
        event = service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()

        continue
      else: raise


    print ('Event created: %s' % (event.get('htmlLink')))

#==============================================================
def ReadAllEvents(service):
    eventsResult = service.events().list(
        calendarId='primary', maxResults=10, singleEvents=True,
        orderBy='startTime').execute()

    events = eventsResult.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        #print(start, event['summary'])
        #print(event['start'].get('timeZone'))
        print(event)
        print("....")


#==============================================================
#read the csv and fill the array
def ReadCSV(GH_file):
  global all_matchs

  #del all_matchs[:]

  with open(GH_file, 'rb') as ghfile:
    r = csv.DictReader(ghfile, delimiter=';', quotechar='"')
    print(r.fieldnames)
    for lines in r:
      #print(lines['competition'])
      if (lines['le']):     # only the ones with a due date
        all_matchs.append(lines)
        
    print("CSV parsing complete...")
      
#==============================================================
def main():
  """Read Gesthand CSV and the entries to Google Calendar.
    """
  try:
    import argparse
    parent = argparse.ArgumentParser(parents=[tools.argparser])
    parent.add_argument("-g","--gesthand_csv", help="Gesthand CSV file", required=True)
    parent.add_argument("-f","--force", help="Force update", required=False)
    flags = parent.parse_args()
    #DBGprint (flags)
    
  except ImportError:
    flags = None
  
  """Creates a Google Calendar API service object and outputs a list of the next
  10 events on the user's calendar.
  """
  credentials = get_credentials(flags)
  http = credentials.authorize(httplib2.Http())
  service = discovery.build('calendar', 'v3', http=http)


  ReadCSV(flags.gesthand_csv)
  StoreMyEvents(service)
  #ReadAllEvents(service)
  print("Complete")

if __name__ == '__main__':
  main()
