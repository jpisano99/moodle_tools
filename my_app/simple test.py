from moodle_box import *
from my_secrets import passwords
import pandas as pd
import json
import time
import os
from settings import app_cfg

# print(json.dumps(results.json(), indent=4, sort_keys=True))

# How to attach lambda to a VPC
# https://www.youtube.com/watch?v=yMzb48BL7qQ&ab_channel=MartyAWSMartyAWS
#


#
# Which a Course / Group to find
#

# course_to_find = 'Secure Workload v3.4 Lab Guide'
# group_to_find = 'CSW-v3.4-TD-Mar-16-2021'
# group_to_find = 'CSW-v3.4-TD-AMER-May-18-2021'

course_to_find = 'Stealthwatch Test Drive'
group_to_find = 'Nov 6 US Test Drive'

# course_to_find = 'Threat Hunting Workshop 4.0'
# group_to_find = '05/27/21 - Prague - Peter Mesjar'

my_sheet = []

#
# Gather all the courses
#
results = mdl_get_courses(passwords['M_TOKEN'], None)
course_dict = results.json()

course_id = ''
category_id = ''
display_name = ''
full_name = ''
for course_record in course_dict:
    if course_record['displayname'] == course_to_find:
        category_id = course_record['categoryid']
        course_id = course_record['id']
        display_name = course_record['displayname']
        full_name = course_record['fullname']
        print('Course FOUND: ', category_id, course_id, display_name, '\t', full_name)
        break

#
# Find all the groups for the selected course
#
results = mdl_get_groups(passwords['M_TOKEN'], course_id)
group_dict = results.json()

group_id = ''
enrollment_key = ''
group_name = ''
for group_record in group_dict:
    if group_record['name'] == group_to_find:
        enrollment_key = group_record['enrolmentkey']
        group_id = group_record['id']
        group_name = group_record['name']

        print('\t Group FOUND: ', group_id, group_name, enrollment_key)
        break

#
# Find all the student profile data and grades for the selected course/group
#
results = mdl_get_group_members(passwords['M_TOKEN'], group_id)
student_dict = results.json()
student_list = student_dict[0]['userids']

for student_id in student_list:
    results = mdl_get_course_user_profile(passwords['M_TOKEN'], student_id, course_id)
    student_profile = results.json()[0]

    student_username = student_profile['username']
    student_firstname = student_profile['firstname']
    student_lastname = student_profile['lastname']
    student_fullname = student_profile['fullname']
    student_email = student_profile['email']
    # student_city = student_profile['city']
    student_country = student_profile['country']
    # print('\t\t', student_id, student_username, student_firstname, student_lastname, student_fullname,
    #      student_email,  student_country)

    student_roles = ''
    for student_role in student_profile['roles']:
        student_roles = student_role['name'] + ' / ' + student_roles
    student_roles = student_roles[:-2]

    #
    # Find all the student grades for the selected course/group
    #
    results = mdl_user_get_grades_items(passwords['M_TOKEN'], course_id, group_id, student_id)
    grade_dict = results.json()
    grade_list = grade_dict['usergrades'][0]['gradeitems']
    lesson_names = []
    lesson_grades = []
    for lesson_grade in grade_list:
        if lesson_grade['itemtype'] == 'mod':
            lesson_names.append(lesson_grade['itemname'])
            lesson_grades.append(lesson_grade['graderaw'])
        elif lesson_grade['itemtype'] == 'course':
            lesson_names.append('Course Total')
            lesson_grades.append(lesson_grade['graderaw'])

    #
    # Make an output row
    #
    my_row = [category_id, course_id, display_name,
              group_id, group_name, enrollment_key,
              student_fullname, student_firstname, student_lastname, student_email,
              student_roles, student_username, student_id, student_country
              ] + lesson_grades
    my_sheet.append(my_row)
    # print(json.dumps(results.json(), indent=4, sort_keys=True))

#
# Create the pandas DataFrame
#
my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
                            app_cfg['WORKING_SUB_DIR'])

my_col_names = ['category_id', 'course_id', 'display_name',
                'group_id', 'group_name', 'enrollment_key',
                'student_fullname', 'student_firstname', 'student_lastname', 'student_email',
                'student_roles', 'student_username', 'student_id', 'student_country',
                ] + lesson_names

df = pd.DataFrame(my_sheet, columns=my_col_names)
my_new_sheet = os.path.join(my_sheet_dir, 'mdl_report.xlsx')
df.to_excel(my_new_sheet, index=False)

exit()



