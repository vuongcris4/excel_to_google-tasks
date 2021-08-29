import pandas as pd #https://www.youtube.com/watch?v=2yMXl5IaIjI https://www.youtube.com/watch?v=7HP5QIoGqpM&list=PL3JVwFmb_BnRYh-UAyaId4Zj6X8zMVNbN&index=2
from Google import Create_Service, convert_to_RFC_datetime

CLIENT_SECRET_FILE = 'client_secret_file.json'
API_NAME = 'tasks'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/tasks']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

pd.set_option("display.max_columns", 100)
pd.set_option("display.max_rows", 500)
pd.set_option("display.min_rows", 500)
pd.set_option("display.max_colwidth", 150)
pd.set_option("display.width", 120)
pd.set_option("display.expand_frame_repr", True)

"""
Insert method
"""
tasklistRestaurant = service.tasklists().insert(
    body = {'title':'Restaurants to cry'}
).execute()

for i in range(97,100):
    service.tasklists().insert(
        body={'title':'Tasklist #{0}'.format(i+1)}
    ).execute()

"""
List method
"""
response = service.tasklists().list().execute()
listItems = response.get('items')
nextPageToken = response.get('nextPageToken')

while nextPageToken:
    response = service.tasklists().list(
        maxResults=30,
        pageToken=nextPageToken
    ).execute()
    listItems.extend(response.get('items'))
    nextPageToken = response.get('nextPageToken')

print(pd.DataFrame(listItems))


"""
Delete method
"""
for item in listItems:
    try:
        if isinstance(int(item.get('title').replace('Tasklist #', '')), int):
            if int(item.get('title').replace('Tasklist #', '')) > 50:
                service.tasklists().delete(tasklist=item.get('id')).execute()
    except:
        pass

response = service.tasklists().list(maxResults=100).execute()
print(pd.DataFrame(response.get('items')))

"""
Update Method
"""
mainTaskList = response.get('items')[6]
mainTaskList['title'] = 'Hihi'
service.tasklists().update(tasklist=mainTaskList['id'], body=mainTaskList).execute()

"""
Get Method
"""
print(service.tasklists().get(tasklist='UjV4Z0RWMWdUeHFGd2VpMA').execute())






# tasklist= "R1FWQkRyWnlDendDXzI4WA"
# service = "https://tasks.googleapis.com/tasks/v1/lists/{tasklist}/tasks"
# tasks = service.tasks().list(tasklist='@default', showHidden=True).execute()  # Modified
# for task in tasks['items']:
#     print(task['title'])
#     print(task['status'])
#     if 'completed' in task:  # Added
#         print(task['completed'])
#     else:
#         print('not completed')
