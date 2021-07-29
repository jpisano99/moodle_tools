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


def merge_sfdc():
    #
    # Gather existing SFDC customer names and file paths
    #
    SFDC_files = os.listdir(SFDC_Repo)

    found_customers = []
    SFDC_file_pathnames = []

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
    # Build a sheet from SFDC Data to Contact email domain data to discover the
    # Cisco official Account_Name

    print('Opening', os.path.join(my_sheet_dir, app_cfg['SFDC_SCRUBBED']))
    df_master = pd.read_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_SCRUBBED']))
    # df_master = pd.read_excel(os.path.join(my_sheet_dir, 'SFDC_Scrubbed_Results_testing.xlsx'))
    print('\tOpened!!!')

    # Create a group_by domain and known as account name
    df_by_domain = df_master.groupby(['domain', 'Account Name'], as_index=False)['Email'].count()

    # Now sort by the most frequently known Account Name
    df_sorted = df_by_domain.sort_values(by=['domain', 'Email'], ascending=[True, False])
    df_sorted .to_excel(os.path.join(my_sheet_dir, app_cfg['SFDC_BY_DOMAIN']), index=False)

    # Clean up
    del df_by_domain
    del df_sorted
    del df_master
    return


if __name__ == "__main__":
    # merge_sfdc()
    # scrub_sfdc()
    build_domain_lookup()



