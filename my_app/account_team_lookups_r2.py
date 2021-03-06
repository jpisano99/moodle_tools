import pandas as pd
import os
from settings import app_cfg
import time
#
# Get Directories and Paths to Files
#
my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
                            app_cfg['ATD_LOOKUPS_SUB_DIR'])


def analyze_registrations():
    # This script scans Cvent Registrations and creates a list of possible search terms
    # to search the ATD
    # It also adds the first TWO "company names" based of the email domain from SFDC data

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
    # Create an interim sheet for analysis
    #
    col_names = ['orig_email',
                 'orig_company_name',
                 'domain',
                 'SFDC Account Names',
                 'search_terms',
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
        domain_name = domain_name.lower()  # Force all domains to lower case

        # Lookup this domain in SFDC for the first two legit Account Name
        sfdc_names = ''
        index = 0
        df_tmp = df_sfdc.loc[df_sfdc['domain'] == domain_name]
        for sfdc_index, sfdc_value in df_tmp.iterrows():
            sfdc_names = sfdc_value['Account Name'] + '::' + sfdc_names
            index += 1
            if index == 2:
                break
        sfdc_names = sfdc_names[:-2]  # Remove trailing field separator (:)
        if len(sfdc_names) == 0:
            sfdc_names = 'None Found'

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

        # Ways to search the ATD for this registrant
        # search_terms = list({scrubbed_name.lower(), first_word.lower(), sld.lower()})
        search_terms = scrubbed_name.lower() + '::' + first_word.lower() + '::' +sld.lower()

        new_row = {'orig_email': email_address,
                   'orig_company_name': company_name,
                   'domain': domain_name,
                   'SFDC Account Names': sfdc_names,
                   'search_terms': search_terms,
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
    # This script creates the file we will give to Power Automate to perform ATD lookups

    #
    # First Inventory the local ATD Directory
    #
    atd_repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
    print('Using Directory:', my_sheet_dir)
    print('ATD Repo Directory:', atd_repo)

    # Gather existing ATD filenames into dict of search_terms
    atd_filenames = os.listdir(atd_repo)
    atd_filename_dict = {}
    print(atd_filenames)

    for filename in atd_filenames:
        if filename[:8] == "ATDData_":
            atd_filename_dict[filename[8:-4]] = os.path.join(atd_repo, filename)

    print('Found: ', len(atd_filenames), ' files in the local ATD repo')

    #
    # Open the registration analysis sheet
    #
    my_domains = os.path.join(my_sheet_dir, app_cfg['REGISTRATION_ANALYSIS'])
    df_domains = pd.read_excel(my_domains)
    print('Opening Registration Analysis Sheet:', my_domains)
    print('\tFound ', len(df_domains), 'customer registrations from Cvent')
    print()

    # Create a group_by domain and known as account name
    # df_by_domain = df_domains.groupby(['domain', 'Account Name'], as_index=False)['Email'].count()
    domains = df_domains['domain'].unique()
    domains.sort()

    # Create a blank sheet for output
    df_atd_lookups = pd.DataFrame(columns=['domain',
                                           'sfdc_account_name',
                                           'search_terms',
                                           'status',
                                           'atd_file_name',
                                           'emails'])
    df_atd_lookups.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_TABLE']), index=False)

    for domain in domains:
        tmp_search_terms = []
        tmp_sfdc_names = ''
        tmp_atd_filename = ''
        tmp_lookup_status = ''
        tmp_emails = []

        filt = (df_domains['domain'] == domain)
        for index, value in df_domains.loc[filt].iterrows():
            tmp_search_terms = tmp_search_terms + value['search_terms'].split('::')
            tmp_sfdc_names = value['SFDC Account Names']
            tmp_emails.append(value['orig_email'])

        # Remove duplicates and sort longest to shortest search terms
        tmp_search_terms = list(set(tmp_search_terms))
        tmp_search_terms.sort(reverse=True, key=len)

        # Format fields with the :: field separator
        tmp_search_terms = '::'.join(map(str, tmp_search_terms))
        tmp_emails = '::'.join(map(str, tmp_emails))

        # Now check the status of this domain in the ATD Repo
        if tmp_sfdc_names == 'None Found':
            tmp_sfdc_names = ''
        tmp_list = tmp_sfdc_names.split('::') + tmp_search_terms.split('::')
        tmp_atd_filename = ''
        tmp_lookup_status = ''
        for x in tmp_list:
            if x in atd_filename_dict:
                tmp_atd_filename = atd_filename_dict[x]
                tmp_lookup_status = "FOUND in Repo via " + x
                break
            else:
                tmp_lookup_status = "NOT FOUND in Repo"

        # if company_name has cisco tag it
        if domain == 'cisco.com':
            # internal_attendees += 1
            tmp_lookup_status = 'DO NOT LOOK UP - Cisco INTERNAL'

        # if domain_name has gmail.com, yahoo.com, outlook.com tag it
        if domain.lower() == 'gmail.com' \
                or domain.lower() == 'yahoo.com' \
                or domain.lower() == 'outlook.com':
            # non_corp_email_attendee += 1
            tmp_lookup_status = 'DO NOT LOOK UP - Non Corporate Email'

        # Add a new row to the output for Power Automate
        data = {'domain': domain,
                'sfdc_account_name': tmp_sfdc_names,
                'search_terms': tmp_search_terms,
                'status': tmp_lookup_status,
                'atd_file_name': tmp_atd_filename,
                'emails': tmp_emails}
        df_atd_lookups = df_atd_lookups.append(data, ignore_index=True)

    df_atd_lookups.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_TABLE']), index=False)

    return

def merge_atd_repo():
    # This pulls all the ATD files in the Repo directory into one master sheet (ATD_RESULTS)
    #
    # Get Directories and Paths to Files
    #
    my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
                                app_cfg['ATD_LOOKUPS_SUB_DIR'])
    atd_repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
    print('Using Directory:', my_sheet_dir)
    print('ATD Repo Directory:', atd_repo)

    #
    # Create a dict of {file: [search_term, domain] to ADD to the ATD lookup files sheets in repo
    #
    ATD_files = os.listdir(atd_repo)
    ATD_filename_dict = {}
    for file in ATD_files:
        if file[:9] == "ATDData__":
            # Extract the search_term and domain from the filename
            tmp_list = file[9:-6].split('__')
            search_term = tmp_list[0]
            domain = tmp_list[1].replace('#', '.')
            ATD_filename_dict[file] = [search_term,domain]

    #
    # Add two columns to each sheet with the domain and search term used in ATD lookups
    #
    for file in ATD_files:
        ATD_file_path = os.path.join(atd_repo, file)
        search_term = ATD_filename_dict[file][0]
        domain = ATD_filename_dict[file][1]
        df_tmp = pd.read_csv(ATD_file_path)
        # See if this file already has the added columns  ?
        if 'search_term' in df_tmp:
            print('Already Updated: ', ATD_file_path)
            continue
        else:
            print('Updating: ', ATD_file_path)
            df_tmp.insert(0,'search_term', search_term)
            df_tmp.insert(0,'domain', domain)
            df_tmp.to_csv(ATD_file_path, index=False)

    #
    # Create a dataframe template for merging all ATD Sheets
    #

    path_name = os.path.join(atd_repo, ATD_files[0])
    df_template = pd.read_csv(path_name)
    df_master = df_template[0:0]
    df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_RESULTS']), index=False)

    print('Processing ', len(ATD_files), ' files in repo')
    for file in ATD_files:
        path_name = os.path.join(atd_repo, file)
        df = pd.read_csv(path_name)
        df_master = pd.concat([df_master, df])

    print('Raw number of Rows', df_master['Customer'].count(), ' processed')

    # Print applymap to escape any unicode characters
    print("Scanning for unicode characters")
    df_master = df_master.applymap(lambda x: x.encode('unicode_escape').
                               decode('utf-8') if isinstance(x, str) else x)

    print("Writing results to ", app_cfg['ATD_RESULTS'] )
    df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_RESULTS']), index=False)
    print()
    print('Total ATD Results', df_master['Customer'].count())
    return


def rename_files():
    #
    # Get Directories and Paths to Files
    #
    # my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
    #                             app_cfg['ATD_LOOKUPS_SUB_DIR'])
    atd_repo = os.path.join(my_sheet_dir, app_cfg['ATD_REPO'])
    print('Using Directory:', my_sheet_dir)
    print('ATD Repo Directory:', atd_repo)

    # Gather existing ATD filenames into dict of search_terms
    atd_filenames = os.listdir(atd_repo)

    tmp = os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_TABLE'])
    df_atd_lookups = pd.read_excel(tmp)

    lookup_dict = {}
    for index, value in df_atd_lookups.iterrows():
        search_terms = ''
        domain = value['domain']
        if isinstance(value['sfdc_account_name'], str):
            search_terms = value['sfdc_account_name'] + "::"

        if isinstance(value['search_terms'], str):
            search_terms = search_terms + value['search_terms']

        search_terms = search_terms.split('::')
        # print(search_terms)
        lookup_dict[domain] = search_terms

    domain = ''
    for file in atd_filenames:
        # Remove the leading and trailing info
        found_by = file[8:-4]

        # for k, v in lookup_dict.items():
        #     if found_by in v:
        #         domain = k

        domain_reformatted = domain.replace('.', '#')
        new_name = 'ATDData__' + found_by + '__' + domain_reformatted + '__.csv'
        old_name = 'ATDData_' + found_by + '.csv'

        new_path = os.path.join(atd_repo, new_name)
        old_path = os.path.join(atd_repo, old_name)
        print (old_path)
        print('\t', new_path)
        print()
        # os.rename(old_path, new_path)

if __name__ == "__main__":
    # analyze_registrations()  # Take Raw Registrations from CVent pages and create additional columns
    # build_atd_lookup_table()  # Take the output of analyze_registrations() and build a file to hand to Power Automate
    merge_atd_repo() # After we have gathered all the ATD lookups MERGE them into one master file
    # rename_files()
