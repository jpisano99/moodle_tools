import pandas as pd
import math
# import xlrd
import os
import time
from settings import app_cfg

#
# Get Directories and Paths to Files
#
my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
                            app_cfg['ATD_LOOKUPS_SUB_DIR'])
ATD_Repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
print('Using Directory:', my_sheet_dir)
print('ATD Repo Directory:', ATD_Repo)

#
# Open the CVent registration sheet
#
my_registration_sheet = os.path.join(my_sheet_dir, app_cfg['RAW_REGISTRATIONS'])
df_registrations = pd.read_excel(my_registration_sheet)
print(df_registrations.shape)
print('Opening Sheet:', my_registration_sheet)

#
# Gather existing ATD customer names and file paths
#
ATD_files = os.listdir(ATD_Repo)
found_customers = []
ATD_file_pathnames = []

for file in ATD_files:
    if file[:8] == "ATDData_":
        ATD_dir_path = os.path.join(ATD_Repo, file)
        ATD_file_pathnames.append(ATD_dir_path)
        found_customers.append(file[8:-4])

#
# Create a sheet for lookups to hand to Power Automate Desktop
#
col_names = ['orig_email',
             'orig_company_name',
             'domain',
             'scrubbed_company_name',
             'first_word_company_name',
             'sld',
             'tld',
             'lookup_status',
             'ATD_filename']
df_status = pd.DataFrame(columns=col_names)
df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_STATUS']), index=False)

#
# Loop over the registration data
#
total_registered = df_registrations['Email Address'].count()
internal_attendees = 0
already_found = 0
domain_dict = {}
search_list = []
lookup_status = ''

for index, value in df_registrations.iterrows():
    lookup_status = ''
    search_list = []

    # Grab the email and company name
    email_address = value['Email Address']
    company_name = value['Company Name']
    try:
        # Is the Company Name a NaN ?
        math.isnan(company_name)
        company_name = "None Specified"
    except:
        pass

    # Do some scrubbing of the name
    scrubbed_name = (company_name.replace(',', '')).strip()  # Replace commas with nulls
    scrubbed_name = (scrubbed_name.replace('.', ' ')).strip()  # Replace Periods with Spaces
    scrubbed_name = (scrubbed_name.replace('\'', '')).strip()  # Replace Apostrophe with nulls
    first_word = (scrubbed_name.split()[0]).strip()  # Get the first word of the name

    # Cut the email in half
    split_email = email_address.split('@')
    user_name = split_email[0]
    domain_name = split_email[1]

    # Check to see if we have already found this company name
    ATD_filename = 'ATDData_' + scrubbed_name + '.csv'
    if os.path.join(ATD_Repo, ATD_filename) in ATD_file_pathnames :
        lookup_status = 'Already FOUND'
        already_found += 1

    # if company_name has cisco skip it
    if 'cisco' in company_name.lower() or domain_name == 'cisco.com':
        internal_attendees += 1
        continue

    # Get the SLD and TLD
    domain_list = domain_name.split('.', 1)
    sld = domain_list[0]
    tld = domain_list[1]

    # Build a dict of search terms by domain name
    # Searching terms are: scrubbed_name, first_word, sld
    # Make everything lowercase for consistency
    if domain_name.lower() in domain_dict:
        data = domain_dict[domain_name.lower()]
        merge_list = data['search_terms'] + [scrubbed_name.lower(), first_word.lower(), sld.lower()]
        merge_list = list(set(merge_list))
        data['search_terms'] = merge_list

        if data['status'] == '':
            data['status'] = lookup_status

        if data['file_name'] == '':
            data['file_name'] = ATD_filename

        domain_dict[domain_name] = data
    else:
        data = {'sld': sld,
                'search_terms': list({scrubbed_name.lower(), first_word.lower(), sld.lower()}),
                'file_name': ATD_filename,
                'status': lookup_status}

        domain_dict[domain_name.lower()] = data

    new_row = {'orig_email': email_address,
               'orig_company_name': company_name,
               'domain': domain_name,
               'scrubbed_company_name': scrubbed_name,
               'first_word_company_name': first_word,
               'sld': sld,
               'tld': tld,
               'lookup_status': lookup_status,
               'ATD_filename': ATD_filename}

    df_status = df_status.append(new_row, ignore_index=True)

# # Make a DataFrame for searches
# domain_dict = {'domain_name': {'sld': ' ',
#                                'search_terms': ['stan', 'blanche'],
#                                'file_name': 'file.txt',
#
my_df = {'domain_name': [],
         'sld': [],
         'search_term': [],
         'status': [],
         'file_name': []
         }

for domain_name, data in domain_dict.items():
    sld = data['sld']
    file_name = data['file_name']
    status = data['status']

    for value in data['search_terms']:
        my_df['domain_name'].append(domain_name)
        my_df['sld'].append(sld)
        my_df['search_term'].append(value)
        my_df['status'].append(status)
        my_df['file_name'].append(file_name)

jim_df = pd.DataFrame(my_df)
print(jim_df)
jim_df.to_excel(os.path.join(my_sheet_dir, 'jim.xlsx'), index=False)
exit()
# df_search = pd.DataFrame(tmp_df)
# print(df_search)
# exit()
# for key, value in domain_dict.items():
#     for tmp_val in value:
#         new_row = {'domain_name': key,
#                   'search_terms': tmp_val,
#                   'file_name': [],
#                   'status': []
#                   }
#         df_search = df_search.append(new_row, ignore_index=True)

# print (df_search)
# # df_search.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_STATUS']), index=False)
# df_search.to_excel(os.path.join(my_sheet_dir, 'jim.xlsx'), index=False)

# Output the Status Sheet
df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_STATUS']), index=False)
print()
print('Total Registered', total_registered)
print('\tInternal Attendees', internal_attendees)
# print('\tDuplicates', duplicates)
print('\tTotal To Be Searched', df_status['orig_email'].count())
print('\t\tAlready Found', already_found)
print('\t\tRemaining Searches to go', df_status['orig_email'].count() - already_found)
exit()

#
# Create a dataframe template for merging all ATD Sheets
#
df_template = pd.read_csv(ATD_file_pathnames[0])
df_master = df_template[0:0]
df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_RESULTS']), index=False)

df_list = []
for file in ATD_file_pathnames:
    df = pd.read_csv(file)
    df_master = pd.concat([df_master, df])
    # print('Num of Rows', df_master['Customer'].count())

df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_RESULTS']), index=False)
print()
print('Total ATD Results', df_master['Customer'].count())
