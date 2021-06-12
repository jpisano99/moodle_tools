import requests
import json
from my_app.my_secrets import passwords


def mdl_get_courses(_api_key, _course_id=None):
    # Web Services Function: core_course_get_courses
    # Arguments
    #   options (Default to "Array ( ) ") options - operator OR is used
    #   List of course id. If empty return all courses

    # REST (POST parameters)
    #   options[ids][0]= int

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_course_get_courses'),
        ('options[ids][0]', _course_id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    course_id = 57  # course_name testing_jim_pisano
    course_id = None
    moodle_courses = mdl_get_courses(m_api_key, course_id)

    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
