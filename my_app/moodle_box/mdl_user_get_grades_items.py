import requests
import json
from my_app.my_secrets import passwords


def mdl_user_get_grades_items(_api_key, _course_id, _group_id=0, _user_id=0):
    # Web Services Function: gradereport_user_get_grade_items
    # Arguments
    #   courseid (Required)
    #   groupid (Default to "0" Get users from this group only
    #   userid (Default to "0") Return grades only for this user (optional)

    # REST (POST parameters)
    #   courseid= int
    #   groupid= int
    #   userid (Default to "0") Return grades only for this user (optional)

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'gradereport_user_get_grade_items'),
        ('courseid', _course_id),
        ('groupid', _group_id),
        ('userid', _user_id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    course_id = 57  # Secure Workload v3.4 Lab Guide
    group_id = 1136  # CSW-v3.4-TD-AMER-May-18-2021
    # user_id = 6152  # Kenny Luke kluke@homestreet.com
    user_id = 4935  # id for JP Cedeno

    moodle_courses = mdl_user_get_grades_items(m_api_key, course_id, group_id, user_id)
    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
