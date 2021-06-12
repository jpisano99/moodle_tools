import requests
import json
from my_app.my_secrets import passwords


def mdl_get_groups(_api_key, _course_id=None):
    # Web Services Function: core_group_get_course_groups
    # Arguments
    #   courseid (Required)

    # REST(POST parameters)
    #   courseid= int

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_group_get_course_groups'),
        ('courseid', _course_id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    # course_id = 64  # course_name testing_jim_pisano
    course_id = 57  # Secure Workload v3.4 Lab Guide
    moodle_courses = mdl_get_groups(m_api_key, course_id)

    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
