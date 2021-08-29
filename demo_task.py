from quickstart import main #https://www.youtube.com/watch?v=_k_oYfKxG8s&list=PL3JVwFmb_BnRYh-UAyaId4Zj6X8zMVNbN&index=3
import pandas as pd
from Google import Create_Service, convert_to_RFC_datetime

CLIENT_SECRET_FILE = 'client_secret_file.json'
API_NAME = 'tasks'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/tasks']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

mainTaskListId = 'UjV4Z0RWMWdUeHFGd2VpMA'

'''
Insert Method
'''

title = 'San Francisco'
notes = ''
due = ''
status = 'needsAction'
deleted = False

request_body = {
    'title': title,
    'notes': notes,
    'due': due,
    'deleted': deleted,
    'status': status
}

response = service.tasks().insert(
    tasklist=mainTaskListId,
    body=request_body
).execute()

responseSanFrancisco = response


def construct_request_body(title, notes=None, due=None, status='needsAction', deleted=False):
    try:
        request_body = {
            'title': title,
            'notes': notes,
            'due': due,
            'deleted': deleted,
            'status': status
        }
        return request_body
    except Exception:
        return None


responseNewYork = service.tasks().insert(
    tasklist=mainTaskListId,
    body=construct_request_body('New York City 2'),
    previous=responseSanFrancisco.get('id')
).execute()


"""
Tasks Demo: Restaurants To Try:
    San Francisco
        - Pearl San Francisco
        - Burma Superstar
        - House of Prime Rib
    Chicago
    New York
"""

service.tasks().insert(
    tasklist=mainTaskListId,
    parent=responseNewYork.get('id'),
    body=construct_request_body(
        'Pearl San Francisco',
        notes='''
        Vuong
        dep
        trai
        ''',
        due=convert_to_RFC_datetime(2021, 8, 18, 20, 30)
    )
).execute()

for i in range(100):
    service.tasks().insert(
        tasklist=mainTaskListId,
        parent=responseSanFrancisco.get('id'),
        body=construct_request_body(
            'Dummy Task #{0}'.format(i+1), due=convert_to_RFC_datetime(2021, (i % 12)+1, )
        )
    ).execute()

'''
List Method
'''
response = service.tasks().list(
    tasklist=mainTaskListId,
    # dueMax=convert_to_RFC_datetime(2021, 8, 17),
    showCompleted=False
).execute()
listItems = response.get('items')
nextPageToken = response.get('nextPageToken')

while nextPageToken:
    response = service.tasks().list(
        tasklist=mainTaskListId,
        # dueMax=convert_to_RFC_datetime(2021, 8, 17),
        showCompleted=False,
        pageToken=nextPageToken
    ).execute()
    listItems.extend(response.get('items'))
    nextPageToken = response.get('nextPageToken')

print(pd.DataFrame(listItems))

'''
Delete Method
'''
for item in listItems:
    service.tasks().delete(tasklist=mainTaskListId, task=item.get('id')).execute()

'''
Update Method
'''
response = service.tasks().list(
    tasklist=mainTaskListId,
    dueMin=convert_to_RFC_datetime(2021, 4, 1),
    dueMax=convert_to_RFC_datetime(2021, 8, 17),
    showCompleted=False
).execute()
listItems = response.get('items')
nextPageToken = response.get('nextPageToken')

while nextPageToken:
    response = service.tasks().list(
        tasklist=mainTaskListId,
        dueMin=convert_to_RFC_datetime(2021, 4, 1),
        dueMax=convert_to_RFC_datetime(2021, 8, 17),
        showCompleted=False,
        pageToken=nextPageToken
    ).execute()
    listItems.extend(response.get('items'))
    nextPageToken = response.get('nextPageToken')

for taskItem in listItems:
    taskItem['status'] = 'completed'
    service.tasks().update(
        tasklist=mainTaskListId,
        task=taskItem.get('id'),
        body=taskItem
    ).execute()

'''
Clear Method
'''
response = service.tasks().list(
    tasklist=mainTaskListId,
    dueMin=convert_to_RFC_datetime(2021, 4, 1),
    dueMax=convert_to_RFC_datetime(2021, 8, 17),
    showCompleted=True
).execute()
listItems = response.get('items')
nextPageToken = response.get('nextPageToken')

while nextPageToken:
    response = service.tasks().list(
        tasklist=mainTaskListId,
        dueMin=convert_to_RFC_datetime(2021, 4, 1),
        dueMax=convert_to_RFC_datetime(2021, 8, 17),
        showCompleted=True,
        pageToken=nextPageToken
    ).execute()
    listItems.extend(response.get('items'))
    nextPageToken = response.get('nextPageToken')

service.tasks().clear(tasklist=mainTaskListId).execute()

'''
Move Method
'''
response = service.tasks().list(
    tasklist=mainTaskListId,
    maxResults=100
).execute()
listItems = response.get('items')
nextPageToken = response.get('nextPageToken')
print(pd.DataFrame(listItems))

service.tasks().move(
    tasklist=mainTaskListId,
    task='TWhBb2l2eWl3ZFZaaElYTg',  # dummy task #10
    parent="d1ZJLWgyeFJyOXBXenBhVQ",  # parent hihi
    previous='MV9IQ1N5dHRmV1FlT3ZIeQ'  # move task #10 bellow (previous) task #9
).execute()
