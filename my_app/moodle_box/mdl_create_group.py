import requests
import json
from my_app.my_secrets import passwords


def mdl_create_group(_api_key, _course_id=None):
    # Web Services Function: core_group_create_groups
    # Arguments
    #   groups (Required) List of group object. A group has a courseid, a name, a description and an enrolment key.

    # REST (POST parameters)
    #   groups[0][courseid]= int
    #   groups[0][name]= string
    #   groups[0][description]= string
    #   groups[0][descriptionformat]= int
    #   groups[0][enrolmentkey]= string
    #   groups[0][idnumber]= string

    # groupname = "stan's group"
    # id_number = 70595
    # enrollment_key = 'abcd1234'
    # description = "test group"

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_group_create_groups'),
        ('groups[0][courseid]', _course_id),
        ('groups[0][name]', groupname),
        ('groups[0][idnumber]', id_number),
        ('groups[0][enrolmentkey]', enrollment_key),
        ('groups[0][description]', description)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    groupname = "stan's group"
    id_number = 70595
    enrollment_key = 'abcd1234'
    description = "test group"

    course_id = 64  # course_name testing_jim_pisano

    moodle_courses = mdl_create_group(m_api_key, course_id)
    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
