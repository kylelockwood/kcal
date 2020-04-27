#! python3

from __future__ import print_function
from datetime import timedelta
import datetime as dt
import os, sys
import pickle
from ics import Calendar, Event
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
SCRIPTPATH = os.path.dirname(os.path.realpath(sys.argv[0])) + '\\'

def create_ics(sourceCalData, outFile='kcal_ics_data', newline=''):
    """ Creates an .ics file for uploading to calendar services """
    cal = Calendar()
    print(f'Writing ICS file \'{outFile}.ics\'... ', flush=True, end='')
    for data in sourceCalData:
        event = Event()
        event.name = data.get('name')
        event.begin = data.get('date')
        event.make_all_day()
        event.description = data.get('description')
        event.location = 'The Oregon Community 700 NE Dekum St. Portland OR'
        cal.events.add(event)
    if not outFile.endswith('.ics'):
        outFile += '.ics'
    with open(outFile, 'w', newline=newline) as f: # Clover calendar wont read ics with extra carriage returns
        f.writelines(cal)
    print('Done')
    return

class gcal():
    def __init__(self, calNames=['primary'], sourceCalData=None):
        self.__service__ = self.__check_creds__(SCRIPTPATH)
        self.cal_ids = self.get_cal_ids()
        self.calNames = calNames
        self.sourceCalData = sourceCalData

    def __check_creds__(self, credPath):
        """
        Opens Google Calendar API connection
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credPath + 'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        service = build('calendar', 'v3', credentials=creds)
        return service

    def get_cal_ids(self):
        """
        Return all calendar names (key) and ids (value) associated with the master calendar email address
        """
        page_token = None
        calids = {}
        while True:
            calendar_list = self.__service__.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                calids[calendar_list_entry['summary']] = calendar_list_entry['id']
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                return calids

    def get_event_ids(self, sourceCalData=None, calNames=None):
        """ Returns a list of event ids of events in sourceCalData """
        # Set default values
        if calNames is None:
            calNames = self.calNames
        if sourceCalData is None:
            sourceCalData = self.sourceCalData
        eventids = []
        # Get the first and last date from the source calendar data
        startDate = dt.datetime(sourceCalData[0]['date'].year, sourceCalData[0]['date'].month, sourceCalData[0]['date'].day).isoformat()+'Z'
        endDate = (dt.datetime(sourceCalData[-1]['date'].year, sourceCalData[-1]['date'].month, sourceCalData[-1]['date'].day)+ timedelta(days=1)).isoformat() + 'Z'
        # Store the dates and names from the source calendar for comparison
        sourceDate = []
        sourceName = []
        for sourceData in sourceCalData:
            sourceDate.append(sourceData['date'].strftime('%Y-%m-%d'))
            sourceName.append(sourceData['name'])
        # Get all events within the dates in the source calendar
        for calName in calNames:
            if calName not in self.cal_ids:
                print(f'Calendar \'{calName}\' not found')
                return None
            for cal in self.cal_ids:
                if  cal == calName:
                    calid = self.cal_ids.get(cal)
                    print(f'Collecting event ids from calendar \'{calName}\'... ', end='', flush=True)
                    events_result = self.__service__.events().list(  calendarId=calid, timeMin=startDate,
                                                            timeMax=endDate, singleEvents=True,
                                                            orderBy='startTime').execute()
                    events = events_result.get('items', [])
                    # if the event name and date match the source calendar name and date, populate event id
                    for eid in events: 
                        idDate = eid['start'].get('date')
                        idName = eid['summary']
                        if idDate in sourceDate and idName in sourceName:
                            eventids.append(eid['id'])
                    print('Done')
        return eventids

    def delete_duplicate_events(self, sourceCalData=None, calNames=None, eventIds='Nothing'):
        """ Deletes events based on eventids """
        # Set default values
        if sourceCalData is None:
            sourceCalData = self.sourceCalData
        if eventIds == 'Nothing':
            eventIds = self.get_event_ids(sourceCalData=sourceCalData)
        if eventIds is None:
            print('No duplicate events found')
            return
        if calNames is None:
            calNames = self.calNames
        # Delete events
        for calName in calNames:
            if calName not in self.cal_ids:
                print(f'Calendar \'{calName}\' not found, no events to delete')
                return
            print('Deleting duplicate events... ', flush=True)
            for cal in self.cal_ids:
                if cal == calName:
                    calid = self.cal_ids.get(cal)
                    for eid in eventIds:
                        print(f'     Deleting event id \'{eid}\'...', end='', flush = True)
                        self.__service__.events().delete(calendarId=calid, eventId=eid).execute()
                        print('Done')
        print('Done')
        return

    def update_gcal(self, sourceCalData=None, calNames=None):
        """
        Uploads event data to the calendars passed
        """
        # Set default values
        if sourceCalData is None:
            sourceCalData = self.sourceCalData
        if calNames is None:
            calNames = self.calNames
        # Create events from sourceCalData
        for calendar in calNames:
            for cal in self.cal_ids:
                if cal == calendar:
                    calid = self.cal_ids.get(cal)
                    print(f'Loading events to calendar \'{cal}\'...')
                    for data in sourceCalData:
                        name = data.get('name')
                        startDate = data.get('date').strftime('%Y-%m-%d')
                        endDate = (data.get('date') + timedelta(days=1)).strftime('%Y-%m-%d')
                        description = data.get('description')
                        location = data.get('location')
                        # Create the event
                        event = {
                            'summary': name,
                            'location': location,
                            'description': description,
                            'start': {
                                'date': startDate,
                                'timeZone': 'America/Los_Angeles',
                            },
                            'end': {
                                'date': endDate,
                                'timeZone': 'America/Los_Angeles',
                            }
                        }
                        event = self.__service__.events().insert(calendarId=calid, body=event).execute()            
                        print(f'     Event \'{name}\' created on {startDate}')
                    print('Done')
            if calendar not in self.cal_ids:
                print(f'Calendar \'{calendar}\' not found.')
        return
