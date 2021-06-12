import requests
import json
from my_app.my_secrets import passwords


def mdl_get_course_user_profile(_api_key, _user_id, _course_id=None):
    # Web Services Function: core_user_get_course_user_profiles
    # Arguments
    #   userlist (Required)

    # REST (POST parameters)
    #   userlist[0][userid]= int
    #   userlist[0][courseid]= int

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_user_get_course_user_profiles'),
        ('userlist[0][userid]', _user_id),
        ('userlist[0][courseid]', _course_id)

        # This works for multiple users on one call
        # ('userlist[0][userid]', '4938'),
        # ('userlist[0][courseid]', '57'),
        # ('userlist[1][userid]', '4935'),
        # ('userlist[1][courseid]', '57')
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    course_id = 57  # Secure Workload v3.4 Lab Guide
    user_id = 4935  # JP Cedeno
    # user_id = 4938  # Yoldy Jacques-Simon
    user_id = 4262

    moodle_courses = mdl_get_course_user_profile(m_api_key, user_id, course_id)

    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
