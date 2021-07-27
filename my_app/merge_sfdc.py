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
SFDC_Repo = os.path.join(my_sheet_dir, app_cfg['SFDC_REPO'])
print('Using Directory:', my_sheet_dir)
print('SFDC Repo Directory:', SFDC_Repo)

#
# Gather existing SFDC customer names and file paths
#
SFDC_files = os.listdir(SFDC_Repo)
found_customers = []
SFDC_file_pathnames = []


def merge_sfdc():
    # #
    # # Get Directories and Paths to Files
    # #
    # my_sheet_dir = os.path.join(app_cfg['HOME'], app_cfg['MOUNT_POINT'], app_cfg['MY_APP_DIR'],
    #                             app_cfg['ATD_LOOKUPS_SUB_DIR'])
    # SFDC_Repo = os.path.join(my_sheet_dir, app_cfg['SFDC_REPO'])
    # print('Using Directory:', my_sheet_dir)
    # print('SFDC Repo Directory:', SFDC_Repo)
    #
    # #
    # # Gather existing SFDC customer names and file paths
    # #
    # SFDC_files = os.listdir(SFDC_Repo)
    # found_customers = []
    # SFDC_file_pathnames = []

    for file in SFDC_files:
        if file[:9] == "SFDCData_":
            SFDC_dir_path = os.path.join(SFDC_Repo, file)
            SFDC_file_pathnames.append(SFDC_dir_path)
            found_customers.append(file[9:-4])

    #
    # Create a dataframe template for merging all ATD Sheets
    #
    print('File being used for Template: ', SFDC_file_pathnames[0])
    df_template = pd.read_excel(SFDC_file_pathnames[3])
    df_master = df_template[0:0]
    df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_RAW']), index=False)

    for file in SFDC_file_pathnames:
        print('Processing: ', file)
        df = pd.read_excel(file)
        df_master = pd.concat([df_master, df])
        print('\tNum of Rows from:', file, df['Account Name'].count())
        print('\tNum of Rows in Master:', df_master['Account Name'].count())
        print()

        # Drop the input dataframe
        del df

    # Write the merge list out and delete dataframes
    df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_RAW']), index=False)
    print()
    print('Total SFDC Results', df_master['Account Name'].count())
    del df_master
    del df_template
    return


def scrub_sfdc():
    print('Opening', os.path.join(my_sheet_dir, app_cfg['SFDC_RAW']))
    df_master = pd.read_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_RAW']))
    print('\tOpened!!!')
    # Add these columns
    df_master["userid"] = ""
    df_master["domain"] = ""
    df_master["sld"] = ""
    df_master["tld"] = ""

    for index, value in df_master.iterrows():
        # Grab the email
        email_address = value['Email']
        # print(email_address, type(email_address))

        # print(value['Account Name'], email_address)
        # print (df_master.loc[index, ['Account Name']])
        if isinstance(email_address, float):
            # print('ERROR ', type(email_address))
            continue

        # Cut the email in half
        split_email = email_address.split('@')
        user_name = split_email[0]
        domain_name = split_email[1]

        # Get the SLD and TLD
        domain_list = domain_name.split('.', 1)
        sld = domain_list[0]
        tld = domain_list[1]

        # Assign the new values
        df_master.loc[index, ['domain']] = [domain_name]
        df_master.loc[index, ['userid']] = [user_name]
        df_master.loc[index, ['sld']] = [sld]
        df_master.loc[index, ['tld']] = [tld]

    print('Writing scrubbed SFDC data')
    df_master.to_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_SCRUBBED']), index=False)
    return


def build_domain_lookup():
    col_names = ['domain',
                 'Account Name'
                 ]
    df_by_domain = pd.DataFrame(columns=col_names)

    print('Opening', os.path.join(my_sheet_dir, app_cfg['SFDC_SCRUBBED']))
    df_master = pd.read_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_SCRUBBED']))
    print('\tOpened!!!')
    domain_dict = {}
    new_row = {}

    for index, value in df_master.iterrows():
        domain = value['domain']
        account_name = value['Account Name']

        # If the domain is blank skip this
        if isinstance(domain, float):
            print('Skipping ', account_name, ' has blank domain')
            continue
        # a-f.ch
        # See if this domain is already in dict
        add_row = False
        if domain not in domain_dict:
            domain_dict[domain] = [account_name]
            add_row = True
        else:
            account_name_list = domain_dict[domain]
            if account_name not in account_name_list:
                account_name_list.append(account_name)
                add_row = True

        if add_row is True:
            new_row = {'domain': domain,
                       'Account Name': account_name
                       }
            df_by_domain = df_by_domain.append(new_row, ignore_index=True)

    df_by_domain.to_excel(os.path.join(my_sheet_dir, 'SFDC_by_domain.xlsx'), index=False)
    return


if __name__ == "__main__":
    build_domain_lookup()
    # scrub_sfdc()
    # merge_sfdc()

