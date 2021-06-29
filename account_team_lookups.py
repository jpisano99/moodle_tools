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
df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)

#
# Loop over the registration data
#
total_registered = df_registrations['Email Address'].count()
internal_attendees = 0
non_corp_email_attendee = 0
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

    # Cut the email in half
    split_email = email_address.split('@')
    user_name = split_email[0]
    domain_name = split_email[1]

    # Get the SLD and TLD
    domain_list = domain_name.split('.', 1)
    sld = domain_list[0]
    tld = domain_list[1]

    try:
        # Is the Company Name a NaN ?
        math.isnan(company_name)
        company_name = sld
    except:
        pass

    # Do some scrubbing of the name
    scrubbed_name = (company_name.replace(',', '')).strip()  # Replace commas with nulls
    scrubbed_name = (scrubbed_name.replace('.', ' ')).strip()  # Replace Periods with Spaces
    scrubbed_name = (scrubbed_name.replace('\'', '')).strip()  # Replace Apostrophe with nulls
    first_word = (scrubbed_name.split()[0]).strip()  # Get the first word of the name

    # Check to see if we have already found this company name
    ATD_filename = 'ATDData_' + scrubbed_name + '.csv'
    if os.path.join(ATD_Repo, ATD_filename) in ATD_file_pathnames :
        lookup_status = 'Already FOUND'
        already_found += 1
    else:
        ATD_filename = ''

    # if company_name has cisco skip it
    if 'cisco' in company_name.lower() or domain_name == 'cisco.com':
        internal_attendees += 1
        continue

    # if domain_name has gmail.com, yahoo.com, outlook.com skip it if company name is blank
    if domain_name.lower() == 'gmail.com' \
       or domain_name.lower() == 'yahoo.com' \
       or domain_name.lower() == 'outlook.com':
        non_corp_email_attendee += 1
        if scrubbed_name == '':
            continue

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

        domain_dict[domain_name.lower()] = data
    else:
        data = {'sld': sld.lower(),
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

#
# END of Main Loop
#


#
# Let's try to find the actual domain name for public email domains
# yahoo.com, gmail.com, outlook.com.....
#
pub_domains = ['yahoo.com', 'gmail.com', 'outlook.com', 'aol.com']
tmp_scrubbed_names = []
for pub_domain in pub_domains:
    if pub_domain in domain_dict:
        tmp_scrubbed_names = domain_dict[pub_domain]['search_terms'] + tmp_scrubbed_names
        del domain_dict[pub_domain]

for x in ['self', 'yahoo', 'gmail', 'outlook', 'aol']:
    try:
        tmp_scrubbed_names.remove(x)
    except:
        pass
tmp_scrubbed_names = list(set(tmp_scrubbed_names))

for scrubbed_name in tmp_scrubbed_names:
    for key, data in domain_dict.items():
        orig_key = key
        if scrubbed_name in data['search_terms']:
            tmp_list = [scrubbed_name] + data['search_terms']
            tmp_list = list(set(tmp_list))

# Sort all the search terms from longest to shortest
# Hopefully we find the MOST specific account team name first
for key, data in domain_dict.items():
    tmp_list= data['search_terms']
    tmp_list.sort(reverse = True, key=len)
    data['search_terms'] = tmp_list

# Make a DataFrame from the clean domain_dict
# This will be for INPUT to Power Automate Desktop
# to use the Account Team Directory Lookup Tool
ATDSearch_data = {'domain_name': [],
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
        ATDSearch_data['domain_name'].append(domain_name)
        ATDSearch_data['sld'].append(sld)
        ATDSearch_data['search_term'].append(value)
        ATDSearch_data['status'].append(status)
        ATDSearch_data['file_name'].append(file_name)

df_ATDSearch = pd.DataFrame(ATDSearch_data)
df_ATDSearch.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_SEARCH_TERMS']), index=False)

# Output the Status Sheet
df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)
print()
print('Total Registered', total_registered)
print('\tInternal Attendees', internal_attendees)
print('\tNon-Corp Email Attendees', non_corp_email_attendee)
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
