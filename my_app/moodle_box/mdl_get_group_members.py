import requests
import json
from my_app.my_secrets import passwords


def mdl_get_group_members(_api_key, _group_id=None):
    # Web Services Function: core_group_get_group_members
    # Arguments groupids (Required)
    #   groupids[0]= int

    # REST (POST parameters)
    #   groupids[0]= int

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_group_get_group_members'),
        ('groupids[0]', _group_id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    group_id = 1137  # Secure Workload v3.4 Lab Guide
    moodle_courses = mdl_get_group_members(m_api_key, group_id)

    print(json.dumps(moodle_courses.json(), indent=4, sort_keys=True))
