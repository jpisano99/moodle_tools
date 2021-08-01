import pandas as pd
import numpy as np
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


def analyze_registrations():

    ATD_Repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
    print('Using Directory:', my_sheet_dir)
    print('ATD Repo Directory:', ATD_Repo)

    #
    # Open the CVent registration sheet
    #
    my_registration_sheet = os.path.join(my_sheet_dir, app_cfg['RAW_REGISTRATIONS'])
    df_registrations = pd.read_excel(my_registration_sheet)
    print('Opening Event Registration Sheet:', my_registration_sheet)
    print('\tFound ', len(df_registrations), 'customer registrations from Cvent')
    print()

    #
    # Open the SFDC by Domain lookup
    #
    my_sfdc_domains_sheet = os.path.join(my_sheet_dir, app_cfg['SFDC_BY_DOMAIN'])
    df_sfdc = pd.read_excel(my_sfdc_domains_sheet)
    print('Opening SFDC Domain Lookups:', my_sfdc_domains_sheet)
    print('\tFound ', len(df_sfdc), 'Company domains in SFDC')
    print()

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
    print('Found: ', len(ATD_file_pathnames), ' files in the local ATD repo')
    print()

    #
    # Create an interim sheet for analysis
    #
    col_names = ['orig_email',
                 'orig_company_name',
                 'domain',
                 'SFDC Account Names',
                 'search_terms',
                 'lookup_status',
                 'ATD_filename',
                 'scrubbed_company_name',
                 'first_word_company_name',
                 'second_level_domain',
                 'top_level_domain'
                 ]
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

        # Lookup this domain in SFDC for the first two legit Account Name
        sfdc_names = ''
        index = 0
        df_tmp = df_sfdc.loc[df_sfdc['domain'] == domain_name]
        for sfdc_index, sfdc_value in df_tmp.iterrows():
            sfdc_names = sfdc_value['Account Name'] + ':' + sfdc_names
            index += 1
            if index == 2:
                break
        sfdc_names = sfdc_names[:-1]  # Remove trailing field separator (:)

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

        # if company_name has cisco tag it
        if 'cisco' in company_name.lower() or domain_name == 'cisco.com':
            internal_attendees += 1
            lookup_status = 'Cisco INTERNAL-DO NOT LOOK UP'

        # if domain_name has gmail.com, yahoo.com, outlook.com tag it
        if domain_name.lower() == 'gmail.com' \
                or domain_name.lower() == 'yahoo.com' \
                or domain_name.lower() == 'outlook.com':
            non_corp_email_attendee += 1
            lookup_status = 'Non Corporate Email-DO NOT LOOK UP'

        # Ways to search the ATD for this registrant
        # search_terms = list({scrubbed_name.lower(), first_word.lower(), sld.lower()})
        search_terms = scrubbed_name.lower() + ':' + first_word.lower() + ':' +sld.lower()

        new_row = {'orig_email': email_address,
                   'orig_company_name': company_name,
                   'domain': domain_name,
                   'SFDC Account Names': sfdc_names,
                   'search_terms': search_terms,
                   'lookup_status': lookup_status,
                   'ATD_filename': ATD_filename,
                   'scrubbed_company_name': scrubbed_name,
                   'first_word_company_name': first_word,
                   'second_level_domain': sld,
                   'top_level_domain': tld
                   }

        df_status = df_status.append(new_row, ignore_index=True)

    print('Unique Domains Found:', len(domain_dict))
    df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS']), index=False)

    return


def build_atd_lookup_table():
    #
    # Open the domain analysis sheet
    #
    my_domains = os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS'])
    df_domains = pd.read_excel(my_domains)
    print('Opening Event Domain Analysis Sheet:', my_domains)
    print('\tFound ', len(df_domains), 'customer registrations from Cvent')
    print()

    # Create a group_by domain and known as account name
    # df_by_domain = df_domains.groupby(['domain', 'Account Name'], as_index=False)['Email'].count()
    domains = df_domains['domain'].unique()
    domains.sort()

    for domain in domains:
        tmp_list = []

        filt = (df_domains['domain'] == domain)
        for index, value in df_domains.loc[filt, ['search_terms', 'SFDC Account Names']].iterrows():
            tmp_list = tmp_list + value['search_terms'].split(':')
            print ('SFDC', value['SFDC Account Names'])

        # Remove duplicates and sort longest to shortest search terms
        tmp_list = list(set(tmp_list))
        tmp_list.sort(reverse=True, key=len)
        print(domain)
        print('\t', tmp_list)
        print()

        # time.sleep(1)
    exit()


    #
    #
    # print('domain', domains)
    #
    # print('jim', domains)
    #print(len(domains), type(domains))
    exit()
    # Now sort by the most frequently known Account Name
    df_sorted = df_by_domain.sort_values(by=['domain', 'Email'], ascending=[True, False])
    df_sorted .to_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_BY_DOMAIN']), index=False)
    domains = {}
    for index, value in df_domains.iterrows():


        pass


    # #
    # # Create a sheet by just domains to hand to Power Automate Desktop
    # #
    col_names = ['domain',
                 'sfdc_account_name',
                 'sfdc_frequency',
                 'search_terms',
                 'lookup_status',
                 'ATD_filename',
                 'email_addresses'
                 ]
    df_by_domain = pd.DataFrame(columns=col_names)
    df_by_domain.to_excel(os.path.join(my_sheet_dir, 'by_domain.xlsx'), index=False)

    domain_name = ''
    # # Build a dict of search terms by domain name
    # # Searching terms are: scrubbed_name, first_word, sld
    # # Make everything lowercase for consistency
    domain_dict = {}
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
                'sfdc_account_name': sfdc_names,
                'emails': [email_address],
                'file_name': ATD_filename,
                'status': lookup_status}

        domain_dict[domain_name.lower()] = data

    return


def merge_ATD_repo():
    # This pulls all the ATD files in the Repo directory into one master sheet (ATD_RESULTS)
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
    analyze_registrations()
    # build_atd_lookup_table()
    # merge_ATD_repo()
