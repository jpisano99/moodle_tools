import json
import requests
import os
import boto3
from my_secrets import passwords
from moodle_box.mdl_get_user_courses import mdl_get_user_courses
from moodle_box.mdl_get_user import mdl_get_user
import pandas as pd

def lambda_handler(event, context):
    print('Event:', event)
    print()
    user_id = event["queryStringParameters"]["user_id"]
    print('Your asking for user_id:', user_id)

    #
    # EFS write a test file
    #
    print('START my_app_data: ', os.listdir("/mnt/my_app_data"))
    with open("/mnt/my_app_data/demofile2.txt", "a") as f:
        f.write("Now the file has more content!")
        f.close()
    print('END my_app_data: ', os.listdir("/mnt/my_app_data"))

    #
    # Make a pandas Dataframe
    #
    my_data = {
        "id": ["4936"],
        "username": ["jonhartwell"],
        "firstname": ["Jon"],
        "lastname": ["Hartwell"]
    }
    df = pd.DataFrame(my_data)
    print(df)
    # df.to_excel("/mnt/my_app_data/jim.xlsx")

    #
    # Make a moodle call
    #
    m_api_key = passwords['M_TOKEN']
    # resp = mdl_get_user_courses(m_api_key, user_id)
    resp = mdl_get_user(m_api_key, user_id)  #
    # my_response = json.dumps(resp.json(), indent=4, sort_keys=True)

    #
    # Write to an S3 Bucket
    #
    bucket_name = "overlook-mountain"
    s3 = boto3.resource('s3')
    s3.Bucket(bucket_name).upload_file('/mnt/my_app_data/jim.csv','jim.csv')

    #
    # Create the Response Object
    #
    responseObject = {}
    responseObject['statusCode'] = 200
    responseObject['headers'] = {}
    responseObject['headers']['Content-Type'] = 'application/json'
    responseObject['body'] = json.dumps(resp.json())

    return responseObject