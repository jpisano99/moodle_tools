import requests
import json
from my_app.my_secrets import passwords


def mdl_get_user(_api_key, _id):
    # Web Services Function: core_user_get_users
    # Arguments
    #   criteria (Required)
    #         the key/value pairs to be considered in user search.
    #         Values can not be empty. Specify different keys only once
    #         (fullname => 'user1', auth => 'manual', ...) - key occurences are forbidden.
    #         The search is executed with AND operator on the criterias. Invalid criterias (keys) are ignored,
    #         the search is still executed on the valid criterias. You can search without criteria,
    #         but the function is not designed for it. It could very slow or timeout.
    #         The function is designed to search some specific users.

    # REST (POST parameters)
    #   criteria[0][key]= string
    #   criteria[0][value]= string

    params = (
        ('moodlewsrestformat', 'json'),
        ('wstoken', _api_key),
        ('wsfunction', 'core_user_get_users'),
        # ('criteria[0][key]', 'lastname'),
        # ('criteria[0][value]', 'luke')
        ('criteria[0][key]', 'id'),
        ('criteria[0][value]', _id)
    )

    resp = requests.get('https://www.ciscosecurityworkshop.com/webservice/rest/server.php', params=params)

    return resp


if __name__ == "__main__":
    # Sample Call and Test Code
    m_api_key = passwords['M_TOKEN']

    # moodle_user = moodle_get_user(m_api_key, 'jpisano99')
    # moodle_user = moodle_get_user(m_api_key, 'radioricki')
    # moodle_user = moodle_get_user(m_api_key, 'jpisano')
    # moodle_user = moodle_get_user(m_api_key, 'majundi')

    moodle_user = mdl_get_user(m_api_key, 4935)  # kluke@homestreet.com

    print(json.dumps(moodle_user.json(), indent=4, sort_keys=True))
