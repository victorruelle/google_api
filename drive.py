import pickle
from googleapiclient.discovery import build
import json
from os import path

CURDIR = path.dirname(path.abspath(__file__))

def create_dir(dir_name,parent_id=None):
    ''' Creates a directory in a parent directory. Duplicates can be created... 
    
    # Input
        - dir_name : str

    # Output
        - dir_id : str

    # TODO : add check if folder exists

    '''
    with open(path.join(CURDIR,"tokens",'token_drive.pickle'), 'rb') as token:
            creds = pickle.load(token)
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {
    'name': dir_name,
    'parents': [parent_id] if parent_id is not None else [],
    'mimeType': 'application/vnd.google-apps.folder'
    }
    f = service.files().create(body=file_metadata,
                                        fields='id').execute()
    return f["id"]


def move_file(file_id,parent_id):
    with open(path.join(CURDIR,"tokens",'token_drive.pickle'), 'rb') as token:
        creds = pickle.load(token)
    service = build('drive', 'v3', credentials=creds)
    response = service.files().get(fileId=file_id,
                                    fields='parents').execute()
    previous_parents = ",".join(response.get('parents'))
    response = service.files().update(fileId=file_id,
                                    addParents=parent_id,
                                    removeParents=previous_parents,
                                    enforceSingleParent=True,
                                    fields='id, parents').execute()
    return response


def share(_id,mail,role="writer",send_notification_email=False):
    '''

    # Input
    - _id : (str) file or directory id
    - mail : (str) email of user to share with
    - role : (str) one of owner, organizer, fileOrganizer, writer, commenter, reader
    - send_notification : (bool) send google's notification email

    '''
    with open(path.join(CURDIR,"tokens",'token_drive.pickle'), 'rb') as token:
        creds = pickle.load(token)
    service = build('drive', 'v3', credentials=creds)
    def callback(request_id, response, exception):
        if exception:
            print(exception)
        else:
            print("Permission Id: %s" % response.get('id'))
    batch = service.new_batch_http_request(callback=callback)
    user_permission = {
        'type': 'user',
        'role': role,
        'emailAddress':mail
    }
    batch.add(service.permissions().create(
            fileId=_id,
            body=user_permission,
            fields='id',
            sendNotificationEmail=send_email
    ))

    batch.execute()




