'''
-----------------------------------------------------------------------
Module python pour importer les dates des matchs dans un agenda google
calendar '''

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from googleapiclient.errors import HttpError

import datetime
from datetime import date, datetime, timedelta
import time
import sys
import getopt
import csv
import codecs
import unicodedata
from scanf import scanf

reload(sys)  
sys.setdefaultencoding('utf8')

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

def semaine(annee, sem):
    ref = date(annee, 1, 4) # Le 4 janvier est toujours en semaine 1
    j = ref.weekday()
    jours = 7*(sem - 1) - j
    lundi = ref + timedelta(days=jours)
    return [lundi + timedelta(days=n) for n in xrange(7)]
#==============================================================
def StoreMyEvents(service):
  for lines in all_matchs:
    ''' This is a dictionnary fill with values
    '''
    event = {}
    event['start']= {}
    event['end']= {}
    event['reminders']= {}

    #fill the values
    # HACK title:
    # Competition is like : "+ 16 masculine secteur 5", "championnat - 11 feminin", championnat regional honneur - 18 ans masculin
    # replace keyword with acronym

    s1 = unicode(lines['competition'], 'utf-8')
    compet = unicodedata.normalize('NFD', s1).encode('ascii', 'ignore')
    print ("\nProcess COMPETITION: {0}".format(compet))
    compet = compet.replace('masculine','M')
    compet = compet.replace('masculin','M')
    compet = compet.replace('feminine','F')
    compet = compet.replace('feminin','F')
    compet = compet.replace('championnat','')
    compet = compet.replace('regional','Reg.')
    compet = compet.replace('honneur','Hon.')

 # Hack Forfait:

#IGNORE    if ((lines['club rec'] == "MONTPELLIER HB 2") or (lines['club vis'] == "MONTPELLIER HB 2")):
#        print( "Skip forfait MONTPELLIER HB 2")
#        continue

    ##event['summary'] = lines['club rec']+"/"+lines['club vis']
    event['summary'] = compet+": "+lines['club rec']+"/"+lines['club vis']
    event['location'] = lines['nom salle']+","+ lines['adresse salle']+","+ lines['CP']+","+ lines['Ville']
    event['description'] = "J"+lines['J']+" "+lines['competition']

    if (lines['le']):
        # Convert date format from dd/mm/yyyy to yyyy-mm-dd
        d = datetime.strptime(lines['le']+" "+lines['horaire'], '%d/%m/%Y %H:%M:%S')
        when = d.strftime('%Y-%m-%dT%H:%M:%S')
        event['start']['dateTime'] = d.strftime('%Y-%m-%dT%H:%M:%S')   #'2017-08-28T09:00:00-07:00'
        event['start']['timeZone'] = 'Europe/Paris'

        #it last 2 hours
        d +=  timedelta(hours=2)
        event['end']['dateTime'] = d.strftime('%Y-%m-%dT%H:%M:%S')
        event['end']['timeZone'] = 'Europe/Paris'
    else:
        week = scanf("%d-%d",lines['\xef\xbb\xbfsemaine'])
        ww=semaine(week[0], week[1])
        print("{0} {1}".format(week, ww))
        print ("GREG: {0}".format(ww[5]))       # Pour le samedi
        # Convert date format from dd/mm/yyyy to yyyy-mm-dd
        # par defaut, 8 heure du mat
        #when = ww.strftime('%Y-%m-%dT8:00:00')
        event['start']['dateTime'] = ww[5].strftime('%Y-%m-%dT8:00:00')   #'2017-08-28T09:00:00-07:00'
        event['start']['timeZone'] = 'Europe/Paris'

        #it last 2 hours
        ww[5] +=  timedelta(hours=2)
        event['end']['dateTime'] = ww[5].strftime('%Y-%m-%dT8:30:00')
        event['end']['timeZone'] = 'Europe/Paris'

        event['summary'] = compet+": "+lines['club rec']+"/"+lines['club vis'] + " !!! HORAIRE PAS ENCORE VALIDE !!!"

        event['reminders']['useDefault'] = False;

    #unique ID !!!
    tag = lines['num poule']+lines['J']
    event['id'] = tag.lower()

    try:
        event['colorId'] = team_list.index(lines['poule'])

    except ValueError:
        team_list.append(lines['poule'])

    #start my colors from 4
    event['colorId'] = 4 + team_list.index(lines['poule'])

    #fix 'id' to be RFC base32 compilant. Should avoid err 400
    for ch in ['v','w', 'x','y','z']:
        if ch in event['id']:
            event['id']=event['id'].replace(ch,"p")
    
    try:

        print (event)
        event = service.events().insert(calendarId='primary', body=event).execute()
        print("---> ADD")

    except HttpError as err:
        if (err.resp.status == 400):
            print("======================================== IGNORE / PLATEAU TO FIX !!!")
            #TODO !!!
            continue
        if (err.resp.status == 403):
            print("========================================time out!!!")
            continue
        if (err.resp.status == 409):  #already exist
            #update the event and retry
            #TODO: si l'event existe deja et que le nouveau n'a pas de date, on skippe le nouveau
            if (lines['le']):
                event = service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()
                continue
            else:
                print ("Skip update")
                continue
        else:
            raise


    print ('Event created: %s' % (event.get('htmlLink')))
    time.sleep(10)

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
      #if (lines['le']):     # only the ones with a due date
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
