import requests
import json
from my_app.my_secrets import passwords


def mdl_get_user_courses(_api_key, _user_id=None):
    # Web Services Function: core_enrol_get_users_courses
    # Arguments
    #   userid (Required)

    # REST (POST parameters)
    #   userid= int

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_enrol_get_users_courses'),
        ('userid', _user_id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    # user_id = 6600  # id for user jpisano99
    # user_id = 5577  # id for Mazen
    # user_id = 56  # id for jpisano
    user_id = 4935  # id for JP Cedeno
    moodle_courses = mdl_get_user_courses(m_api_key, user_id)

    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
