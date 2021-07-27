import pandas as pd
import math
# import xlrd
import os
import time
from settings import app_cfg


def analyze_registrations():
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
    # Create an interim sheet for analysis
    #
    col_names = ['orig_email',
                 'orig_company_name',
                 'domain',
                 'scrubbed_company_name',
                 'first_word_company_name',
                 'second_level_domain',
                 'top_level_domain',
                 'lookup_status',
                 'ATD_filename',
                 'search_terms',
                 'email_addresses',
                 'SFDC Account Name'
                 ]
    df_status = pd.DataFrame(columns=col_names)
    df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)

    #
    # Create a sheet by just domains to hand to Power Automate Desktop
    #
    col_names = ['domain',
                 'search_terms',
                 'lookup_status',
                 'ATD_filename',
                 'email_addresses'
                 ]
    df_by_domain = pd.DataFrame(columns=col_names)
    df_by_domain.to_excel(os.path.join(my_sheet_dir, 'by_domain.xlsx'), index=False)

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

        # If there is NO company name use the sld as the company name
        if not isinstance(company_name, str):
            company_name = sld

        # Do some scrubbing of the name
        scrubbed_name = (company_name.replace(',', '')).strip()  # Replace commas with nulls
        scrubbed_name = (scrubbed_name.replace('.', ' ')).strip()  # Replace Periods with Spaces
        scrubbed_name = (scrubbed_name.replace('\'', '')).strip()  # Replace Apostrophe with nulls
        first_word = (scrubbed_name.split()[0]).strip()  # Get the first word of the name

        # Check to see if we have already found this company name
        ATD_filename = 'ATDData_' + scrubbed_name + '.csv'
        if os.path.join(ATD_Repo, ATD_filename) in ATD_file_pathnames:
            lookup_status = 'Already FOUND'
            already_found += 1
        else:
            ATD_filename = ''

        # if company_name has cisco skip it
        if 'cisco' in company_name.lower() or domain_name == 'cisco.com':
            internal_attendees += 1
            lookup_status = 'Cisco INTERNAL-DO NOT LOOK UP'

        # if domain_name has gmail.com, yahoo.com, outlook.com skip it if company name is blank
        if domain_name.lower() == 'gmail.com' \
                or domain_name.lower() == 'yahoo.com' \
                or domain_name.lower() == 'outlook.com':
            non_corp_email_attendee += 1
            lookup_status = 'Non Corporate Email-DO NOT LOOK UP'

        # Build a dict of search terms by domain name
        # Searching terms are: scrubbed_name, first_word, sld
        # Make everything lowercase for consistency
        if domain_name.lower() in domain_dict:
            data = domain_dict[domain_name.lower()]

            # updates the list of search terms
            merge_list = data['search_terms'] + [scrubbed_name.lower(), first_word.lower(), sld.lower()]
            merge_list = list(set(merge_list))  # Remove duplicate search terms
            merge_list.sort(key=len, reverse=True)  # Sort the list from longest to shortest search term
            data['search_terms'] = merge_list

            # Keep a list of email addresses for this domain
            merge_list = data['emails'] + [email_address]
            merge_list = list(set(merge_list))
            data['emails'] = merge_list

            if data['status'] == '':
                data['status'] = lookup_status

            if data['file_name'] == '':
                data['file_name'] = ATD_filename

            domain_dict[domain_name.lower()] = data
        else:
            # This is the first occurrence of this domain

            search_terms = list({scrubbed_name.lower(), first_word.lower(), sld.lower()})
            search_terms.sort(key=len, reverse=True)  # Sort the list from longest to shortest search term

            data = {'sld': sld.lower(),
                    'search_terms': search_terms,
                    'emails': [email_address],
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
                   'ATD_filename': ATD_filename,
                   'search_terms': data['search_terms']}

        df_status = df_status.append(new_row, ignore_index=True)
    #
    # END of Main Loop
    #

    #
    # Now loop over by_domain to build the Desktop Automate List
    #
    for domain, data in domain_dict.items():
        for search_term in data['search_terms']:
            # Check if we have found this search term and it's in the repo
            # Make everything lowercase for this test
            ATD_filename = 'ATDData_' + search_term + '.csv'
            lookup_status = ''
            if (os.path.join(ATD_Repo, ATD_filename)).lower() in (string.lower() for string in ATD_file_pathnames):
                lookup_status = 'Already FOUND'
                already_found += 1
            else:
                ATD_filename = ''

            new_row = {'domain': domain,
                       'search_terms': search_term,
                       'lookup_status': lookup_status,
                       'ATD_filename': ATD_filename,
                       'email_addresses': data['emails']
                       }
            df_by_domain = df_by_domain.append(new_row, ignore_index=True)

    df_by_domain.to_excel(os.path.join(my_sheet_dir, 'by_domain.xlsx'), index=False)

    print('Unique Domains Found:', len(domain_dict))
    df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)

    return


def merge_ATD_repo():
    #
    # Get Directories and Paths to Files
    #
    my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
                                app_cfg['ATD_LOOKUPS_SUB_DIR'])
    ATD_Repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
    print('Using Directory:', my_sheet_dir)
    print('ATD Repo Directory:', ATD_Repo)

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
    return


if __name__ == "__main__":
    # analyze_registrations()
    merge_ATD_repo()



#
# #
# # Let's try to find the actual domain name for public email domains
# # yahoo.com, gmail.com, outlook.com.....
# #
# pub_domains = ['yahoo.com', 'gmail.com', 'outlook.com', 'aol.com']
# tmp_scrubbed_names = []
# for pub_domain in pub_domains:
#     if pub_domain in domain_dict:
#         tmp_scrubbed_names = domain_dict[pub_domain]['search_terms'] + tmp_scrubbed_names
#         del domain_dict[pub_domain]
#
# for x in ['self', 'yahoo', 'gmail', 'outlook', 'aol']:
#     try:
#         tmp_scrubbed_names.remove(x)
#     except NameError:
#         print('Found error')
# tmp_scrubbed_names = list(set(tmp_scrubbed_names))
#
# for scrubbed_name in tmp_scrubbed_names:
#     for key, data in domain_dict.items():
#         orig_key = key
#         if scrubbed_name in data['search_terms']:
#             tmp_list = [scrubbed_name] + data['search_terms']
#             tmp_list = list(set(tmp_list))
#
# # Sort all the search terms from longest to shortest
# # Hopefully we find the MOST specific account team name first
# for key, data in domain_dict.items():
#     tmp_list = data['search_terms']
#     tmp_list.sort(reverse=True, key=len)
#     data['search_terms'] = tmp_list
#
# # Make a DataFrame from the clean domain_dict
# # This will be for INPUT to Power Automate Desktop
# # to use the Account Team Directory Lookup Tool
# ATDSearch_data = {'domain_name': [],
#                   'sld': [],
#                   'search_term': [],
#                   'status': [],
#                   'file_name': []
#                   }
#
# for domain_name, data in domain_dict.items():
#     sld = data['sld']
#     file_name = data['file_name']
#     status = data['status']
#
#     for value in data['search_terms']:
#         ATDSearch_data['domain_name'].append(domain_name)
#         ATDSearch_data['sld'].append(sld)
#         ATDSearch_data['search_term'].append(value)
#         ATDSearch_data['status'].append(status)
#         ATDSearch_data['file_name'].append(file_name)
#
# df_ATDSearch = pd.DataFrame(ATDSearch_data)
# df_ATDSearch.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_SEARCH_TERMS']), index=False)
#
# # Output the Status Sheet
# df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)
# print()
# print('Total Registered', total_registered)
# print('\tInternal Attendees', internal_attendees)
# print('\tNon-Corp Email Attendees', non_corp_email_attendee)
# print('\tTotal To Be Searched', df_status['orig_email'].count())
# print('\t\tAlready Found', already_found)
# print('\t\tRemaining Searches to go', df_status['orig_email'].count() - already_found)
#
# exit()
