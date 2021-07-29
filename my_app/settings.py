from my_app.my_secrets import passwords
import os


# application predefined constants
app_cfg = dict(
    # RUNTIME_ENV='AWS',
    RUNTIME_ENV='LOCAL',
    # RUNTIME_ENV='PYTHONANYWHERE',

    VERSION=1.0,
    GITHUB="{url}",
    HOME=os.path.expanduser("~"),
    MOUNT_POINT='my_app_data',
    MY_APP_DIR='test_drives',
    WORKING_SUB_DIR='cvent_registrations',
    UPDATES_SUB_DIR='',
    ARCHIVES_SUB_DIR='',
    PROD_DATE='',
    UPDATE_DATE='',
    META_DATA_FILE='config_data.json',

    # Cvent Registration Directories and FilesRaw Data Files to ingest
    RAW_REGISTRATIONS='Secure_Workload_Tetration_as_of_6_7_21.xlsx',
    REGISTRATION_ANALYSIS='Registration_Analysis.xlsx',

    # Scrubbed Data Working Files
    XLS_SUBSCRIPTIONS='tmp_Master Subscriptions.xlsx',

    # SFDC Directories and Files
    SFDC_REPO='SFDC_Files',
    SFDC_RAW='SFDC_Raw_Results.xlsx',
    SFDC_SCRUBBED='SFDC_Scrubbed_Results.xlsx',
    SFDC_BY_DOMAIN='SFDC_by_domain.xlsx',

    # Account Team Directories (ATD) and Files
    ATD_LOOKUPS_SUB_DIR='AccountTeamLookups',
    ATD_REPO='ATD_Files',
    ATD_SEARCH_TERMS='ATD_Search_Terms.xlsx',
    ATD_RESULTS='ATD_Raw_Results.xlsx',





    #
    # Testing
    #

    XLS_UNIQUE_CUSTOMERS='',

    # SmartSheet Sheets and Names
    SS_SAAS='SaaS customer tracking',
)

# database configuration settings
# selected via the app_cfg['RUNTIME_ENV'] setting
if app_cfg['RUNTIME_ENV'] == 'AWS':
    # This is for a AWS based SQL db
    db_config = dict(
        DATABASE="test_db",
        USER="admin",
        PASSWORD=passwords["DB_PASSWORD"],
        HOST="database-1.cp1kaaiuayns.us-east-1.rds.amazonaws.com"
    )
elif app_cfg['RUNTIME_ENV'] == 'PYTHONANYWHERE':
    # This is for PythonAnywhere based SQL db
    db_config = dict(
        DATABASE="jpisano$test_db",
        USER="jpisano",
        PASSWORD=passwords["DB_PASSWORD"],
        HOST="jpisano.mysql.pythonanywhere-services.com"
    )
elif app_cfg['RUNTIME_ENV'] == 'LOCAL':
    # This is for a local based SQL db
    db_config = dict(
        DATABASE="ta_adoption_db",
        USER="root",
        PASSWORD=passwords["DB_PASSWORD"],
        HOST="localhost"
    )

# Smart sheet Config settings
ss_token = dict(
    SS_TOKEN=passwords["SS_TOKEN"]
)

