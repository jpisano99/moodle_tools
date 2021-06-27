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
print('Using Directory:', my_sheet_dir)

#
# Open the CVent registration sheet
#
my_registration_sheet = os.path.join(my_sheet_dir, app_cfg['RAW_REGISTRATIONS'])
df_registrations = pd.read_excel(my_registration_sheet)
print(df_registrations.shape)
print('Opening Sheet:', my_registration_sheet)

#
# Create a sheet for lookups to hand to Power Automate Desktop
#
col_names= ['orig_email',
            'orig_company_name',
            'domain',
            'scrubbed_company_name',
            'first_word_company_name',
            'sld',
            'tld',
            'filename']
df_status = pd.DataFrame(columns=col_names)
df_status.to_excel(os.path.join(my_sheet_dir, app_cfg['ATD_LOOKUP_STATUS']), index=False)

#
# Loop over the registration data
#
# print(df_registrations[['Email Address', 'Company Name']])
total_registered = df_registrations['Email Address'].count()
internal_attendees = 0
duplicates = 0
unique_searches = []
for index, value in df_registrations.iterrows():
    # Grab the email and company name
    email_address = value['Email Address']
    company_name = value['Company Name']
    ATD_filename = 'ATDData_' + company_name + '.csv'

    try:
        math.isnan(company_name)
        company_name = "None Specified"
    except:
        pass

    # Cut the email in half
    split_email = email_address.split('@')
    user_name = split_email[0]
    domain_name = split_email[1]

    # if company_name has cisco skip it
    if 'cisco' in company_name.lower() or domain_name == 'cisco.com':
        # print('found', company_name)
        internal_attendees += 1
        continue

    if company_name+domain_name in unique_searches:
        # print('found duplicate ', company_name, '/ ', domain_name)
        duplicates += 1
        continue
    else:
        unique_searches.append(company_name+domain_name)

    # Get the SLD and TLD
    domain_list = domain_name.split('.', 1)
    sld = domain_list[0]
    tld = domain_list[1]

    # print('emai address', email_address)
    # print('\tuser name', user_name)
    # print('\tcompany name', company_name)
    # print('\tdomain name', domain_name)
    # print('\tsld name', sld)
    # print('\ttld name', tld)
    # print()
    # time.sleep(1)

    new_row = {'orig_email': email_address,
               'orig_company_name': company_name,
               'domain': domain_name,
               'scrubbed_company_name': '',
               'first_word_company_name': '',
               'sld': sld,
               'tld': tld
               'ATD_filename': ATD_filename}

    df_status=df_status.append(new_row, ignore_index=True)

df_status.to_excel(os.path.join(my_sheet_dir, 'blanche.xlsx'))
print()
print('Total Registered', total_registered)
print('\tInternal Attendees', internal_attendees)
print('\tDuplicates', duplicates)
print('\tTotal To Be Searched', df_status['orig_email'].count())

exit()


#
# Gather existing ATD customer names and file paths
#
files = os.listdir(my_sheet_dir)
found_customers = []
atd_filenames = []

for file in files:
    if file[:8] == "ATDData_":
        file_path = os.path.join(my_sheet_dir, file)
        atd_filenames.append(file_path)
        found_customers.append(file[8:-4])
# print(found_customers)
# print(atd_filenames)
exit()






#
# Create a dataframe template for merging all ATD Sheets
#
df_template = pd.read_csv(atd_filenames[0])
df_master = df_template[0:0]
df_master.to_excel(os.path.join(my_sheet_dir, 'Master_Account_Teams.xlsx'))

df_list = []
# new = pd.DataFrame([old.A, old.B, old.C]).transpose()
# result = pd.concat(frames)
for file in atd_filenames:
    df = pd.read_csv(file)
    # print(df)
    # df_list.append(df)
    df_master = pd.concat([df_master, df])
    print('Num of Rows', df_master['Customer'].count())
    # time.sleep(1)

df_master.to_excel(os.path.join(my_sheet_dir, 'Master_Account_Teams.xlsx'), index=False)

#
# df_registrations = pd.read_excel(my_sheet)
# print(df_registrations)


# Lookup by name provided (Alpha Characters only)
# Lookup by name provided (First Word only)

# Remove
# Lookup by email
#
