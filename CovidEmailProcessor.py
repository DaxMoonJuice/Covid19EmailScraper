import json
import re
import sys
from datetime import datetime

import win32com.client
import logging
from pandas import DataFrame, isna, concat


def regex_extractor(extraction_rules):
    extracted_dict = {}
    for value_to_extract, extraction_rule in extraction_rules.items():

        result = re.search(extraction_rule['regex'], extraction_rule['text'], re.MULTILINE)
        if type(result) is re.Match:
            extracted_dict[value_to_extract] = result.group()

        if result is None:
            logging.debug(
                {"message": "Could Not Extract Value From Email - Exiting Value Extraction",
                 "value_to_extract": value_to_extract, "text_used":
                     extraction_rule['text']})
            return None

    return extracted_dict


class EmailFactory():
    def __init__(self):
        self.__map = {"TestRegistrationEmail": TestRegistrationEmail,
                      "PositivePCRTestResultEmail": PositivePCRTestResultEmail,
                      "NegativeLateralFlowTestResultEmail": NegativeLateralFlowTestResultEmail,
                      "NegativePCRTestResultEmail": NegativePCRTestResultEmail}

    def generate(self, object_type, outlook_email_obj):
        return self.__map[object_type](outlook_email_obj)


class TestRegistrationEmail():
    def __init__(self, outlook_email_obj):
        self.__outlook_email_obj = outlook_email_obj

    def identify(self, email_body_text):
        has_barcode_ref = re.search(r'Test kit barcode reference: .*', email_body_text)
        has_confirmation_msg = re.search(r'Kit registration confirmed', email_body_text)

        if has_barcode_ref is None:
            return False

        if has_confirmation_msg is None:
            return False

        else:
            return True

    def extract_values(self, *args, **kwargs):

        regex_extraction_rules = {
            "name": {"regex": ".*(?=\r\n\r\n\r\nKit registration confirmed)", "text": self.__outlook_email_obj.body},
            "pcr_test_kit_barcode_ref": {"regex": r"(?<=Test kit barcode reference: ).*(?=\r\n)",
                                         "text": self.__outlook_email_obj.body},
        }

        extracted_fields = regex_extractor(regex_extraction_rules)

        if extracted_fields is None:
            return None

        extracted_fields["date_email_received"] = str(self.__outlook_email_obj.SentOn)
        extracted_fields["email_type"] = "TestRegistrationEmail"
        extracted_fields["email_subject"] = self.__outlook_email_obj.Subject
        extracted_fields["test_type"] = "PCR"

        return extracted_fields


class NegativeLateralFlowTestResultEmail():
    def __init__(self, outlook_email_obj):
        self.__outlook_email_obj = outlook_email_obj

    def identify(self, email_body_text):
        has_negative_test_message = re.search(
            r'Your coronavirus lateral flow test result is negative. It’s likely you were not infectious when the test was done',
            email_body_text)

        if has_negative_test_message is None:
            return False

        else:
            return True

    def extract_values(self, *args, **kwargs):
        regex_extraction_rules = {
            "name": {"regex": r"(?<=Dear ).* .*(?=\r\n)", "text": self.__outlook_email_obj.body},
            "test_date": {"regex": r"(?<=Test date: )\d\d .* \d\d\d\d(?=\r\n)", "text": self.__outlook_email_obj.body},
        }

        extracted_fields = regex_extractor(regex_extraction_rules)
        if extracted_fields is None:
            return None

        extracted_fields["date_email_received"] = str(self.__outlook_email_obj.SentOn)
        extracted_fields["result"] = "negative"
        extracted_fields["email_type"] = "NegativeLateralFlowTestResultEmail"
        extracted_fields["email_subject"] = self.__outlook_email_obj.Subject
        extracted_fields["test_type"] = "Lateral Flow Test"

        return extracted_fields


class PositivePCRTestResultEmail():
    def __init__(self, outlook_email_obj):
        self.__outlook_email_obj = outlook_email_obj

    def identify(self, email_body_text):
        has_positive_test_message = re.search(
            r'Your recent coronavirus test has come back positive\.|Your coronavirus PCR test \(or other lab test\) result is positive\. It’s likely you had the virus when the test was done\.',
            email_body_text)

        if has_positive_test_message is None:
            return False

        else:
            return True

    def extract_values(self, *args, **kwargs):
        regex_extraction_rules = {
            "name": {"regex": r"(?<=Dear ).* .*(?=\r\n)|(?<=Hello ).* .*(?=\r\n)",
                     "text": self.__outlook_email_obj.body},
            "test_date": {"regex": r"(?<=Test date: )\d\d .* \d\d\d\d(?=\r\n)", "text": self.__outlook_email_obj.body},
        }

        extracted_fields = regex_extractor(regex_extraction_rules)
        if extracted_fields is None:
            return None

        extracted_fields["date_email_received"] = str(self.__outlook_email_obj.SentOn)
        extracted_fields["result"] = "positive"
        extracted_fields["email_type"] = "PositivePCRTestResultEmail"
        extracted_fields["email_subject"] = self.__outlook_email_obj.Subject
        extracted_fields["test_type"] = "PCR"

        return extracted_fields


class NegativePCRTestResultEmail():
    def __init__(self, outlook_email_obj):
        self.__outlook_email_obj = outlook_email_obj

    def identify(self, email_body_text):
        has_negative_test_message = re.search(
            r'Your coronavirus test result is negative\. It’s likely you did not have the virus when the test was done|Your recent coronavirus test has come back negative',
            email_body_text)

        if has_negative_test_message is None:
            return False

        else:
            return True

    def extract_values(self, *args, **kwargs):

        regex_extraction_rules = {
            "name": {"regex": r"(?<=Dear ).* .*(?=\r\n)", "text": self.__outlook_email_obj.body},
            "test_date": {"regex": r"(?<=Test date: )\d\d .* \d\d\d\d(?=\r\n)", "text": self.__outlook_email_obj.body},
        }

        extracted_fields = regex_extractor(regex_extraction_rules)
        if extracted_fields is None:
            return None

        extracted_fields["date_email_received"] = str(self.__outlook_email_obj.SentOn)
        extracted_fields["result"] = "negative"
        extracted_fields["email_type"] = "NegativePCRTestResultEmail"
        extracted_fields["email_subject"] = self.__outlook_email_obj.Subject
        extracted_fields["test_type"] = "PCR"

        return extracted_fields


def construct_email_template_object(email_object):
    classesToIdentify = [TestRegistrationEmail, PositivePCRTestResultEmail, NegativeLateralFlowTestResultEmail,
                         NegativePCRTestResultEmail]

    identification_results = {}

    for email_class in classesToIdentify:
        identification_results[f'{email_class.__name__}'] = email_class.identify(email_class, email_object.Body)

    succesfull_matches = []
    failed_matches = []
    for key, value in identification_results.items():

        if value == True:
            succesfull_matches.append(key)
        else:
            failed_matches.append(key)

    if len(succesfull_matches) > 1:
        logging.debug(
            {"msg": "Two Or More Templates Where Matched To Email", "identification_results": identification_results,
             "email_body": email_object.Body})
        return None

    if len(succesfull_matches) == 0:
        return None
    else:
        return EmailFactory().generate(object_type=succesfull_matches[0], outlook_email_obj=email_object)


def extract_account(email_addr, mapi):
    for account in mapi.Accounts:
        if email_addr == account.DisplayName:
            return account

    return None


def extract_folder(mapi, folder_path, account_name, _folder_obj=None):
    # split the path
    folders = folder_path.split('/')

    # if the folder does not yet exist we fetch it from the the MAPI
    if _folder_obj == None:
        _folder_obj = mapi.Folders(account_name).Folders(folders[0])

    # if the folders path only has 1 item we've hit the endpoint can return results!
    if len(folders) == 1:
        return _folder_obj

    # otherwise call again with the next item in the path
    new_folder_path = '/'.join(folders[1:])
    _folder_obj = _folder_obj.Folders(folders[1])

    return extract_folder(mapi, new_folder_path, account_name, _folder_obj)


def extract_emails(folder_obj):
    return list(folder_obj.Items)


def datestamp_to_datetime(x, input_sftime):
    if isna(x):
        return x
    datetime_obj = datetime.strptime(x, input_sftime)
    return datetime_obj


def generate_week_num(x):
    if isna(x):
        return x

    week_num = x.isocalendar()[1]
    return week_num


def log_stats(results_df):
    logging.info(
        f" \n \n----Execution Has Completed---\nSuccesfully Extracted Results From {results_df.shape[0]} Emails \nResults Have Been Written To A CSV File Named 'dumped_results.csv\n")


def read_site_file(file_path):
    try:
        json_file = open(file_path, mode='r')
    except FileNotFoundError:
        logging.error(f"FileNotFoundError: Could Not Find File Named '{file_path}'")
        return None
    try:
        site_object = json.load(json_file)

    except json.decoder.JSONDecodeError:
        logging.error(f"JSONDecodeError: Encountered An Error Trying To Parse JSON In The File Named '{file_path}'")
        return None

    return site_object


def validate_site_feed(site_list):
    required_keys = ["Site_Name", "Email_Account", "Folder_Path"]

    for site in site_list:
        # Are required keys present
        for key in site.keys():
            if key not in required_keys:
                logging.error({"message": f"SiteListValidationError - JSON Key Was Not Allowed", "JSON Key Not Allowed": key,
                               "JSON Keys That Are Allowed": required_keys, "Site Where Error Was Seen": site})
                return None

        # does the Email_Account field contain an email address - email match regex pulled from here https://www.codegrepper.com/code-examples/whatever/regex+to+identify+email+address
        has_email=re.search(r"^[a-zA-Z0-9.!#$%&’*+\/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$",site["Email_Account"])
        if (has_email) is None:
            logging.error({"message": f"SiteListValidationError - Value of Email_Account Is Not An Email Address", "JSON Key": "Email_Account",
                           "Invalid Valid": site["Email_Account"], "Site Where Error Was Seen": site})


    return site_list


def run_script(dev):

    if dev == True:
        raw_test_sites = read_site_file("site_list_dev.json")
    else:
        raw_test_sites = read_site_file("site_list_prod.json")


    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    list_of_dataframes = []

    if raw_test_sites == None:
        return None

    else:
        test_sites = validate_site_feed(raw_test_sites)

    for test_site in test_sites:
        email_account = test_site["Email_Account"]
        folder_path = test_site["Folder_Path"]
        site_name = test_site["Site_Name"]

        # extract account for folder object

        outlook_account = extract_account(email_account, mapi)
        if outlook_account is None:
            logging.warning(f"Could Not Find Outlook Account For {email_account} - Skipping Account")
            continue

        logging.info(
            f"\n\n ----Fetching Emails For Account '{email_account}' In Folder '{folder_path}' For Site Named '{site_name}'---\n")

        outlook_folder = extract_folder(mapi, folder_path, outlook_account.DisplayName)

        email_list = extract_emails(outlook_folder)

        email_contents = []

        if email_list == []:
            logging.warning(
                f"No Emails Exist For For Account '{email_account}' From Folder '{folder_path}' - Skipping Folder")
            continue

        logging.info(f"Fetched {len(email_list)} Emails For Account '{email_account}' From Folder '{folder_path}'")
        logging.info(
            f"Extracting Values From {len(email_list)} Emails For Account '{email_account}' From Folder '{folder_path}'")
        for email in email_list:
            email_template = construct_email_template_object(email)

            if email_template is None:
                logging.warning(
                    f'Could Not Match An Email_Template For Email With Subject "{email.Subject}" Received At {str(email.ReceivedTime)} - Skipping Email')
                continue

            logging.debug(f"Email With Subject '{email.Subject}' is an instance of {type(email_template)}")

            extracted_values = email_template.extract_values()
            extracted_values["site_name"] = site_name

            if extracted_values is None:
                logging.warning(
                    f"Could Not Extract Values For Email With Subject '{email.Subject}' And Type Of {type(email_template)}")
                continue

            email_contents.append(extracted_values)
            logging.debug(
                f"Extracted Contents Of Email With Subject '{email.Subject}' And Type Of {type(email_template)}")

        logging.info(f"Successfully Extracted Values For {len(email_contents)} Emails")
        logging.warning(f'Failed To Extract Values For {len(email_list) - len(email_contents)} Emails')

        email_contents_df = DataFrame(email_contents)

        email_contents_df['test_date'] = email_contents_df['test_date'].apply(datestamp_to_datetime,
                                                                              args=(["%d %B %Y"]))

        logging.debug("Converted test_date col to datetime")
        email_contents_df['date_email_received'] = email_contents_df['date_email_received'].apply(
            lambda x: re.sub(r'(?<=\+\d\d):', '', x))

        logging.debug("Fixed timezone stamp for date_email_received column")
        email_contents_df['date_email_received'] = email_contents_df['date_email_received'].apply(datestamp_to_datetime,
                                                                                                  args=([
                                                                                                      '%Y-%m-%d %H:%M:%S%z']))

        logging.debug("Converted date_email_received col to datetime")

        email_contents_df['email_received_in_week_num'] = email_contents_df['date_email_received'].apply(
            generate_week_num)

        logging.debug("Calculated Calendar Week Number For date_email_received column")
        email_contents_df['test_date_occured_in_week_num'] = email_contents_df['test_date'].apply(
            generate_week_num)

        logging.debug("Calculated Calendar Week Number For test_date_occured_in_week_num column")

        # export to csv

        email_contents_df.rename(columns={'name': 'Testee Name', 'pcr_test_kit_barcode_ref': 'PCR Test Kit Barcode',
                                          'date_email_received': 'Date Email Received',
                                          'email_subject': 'Email Subject', 'site_name': 'Site Name',
                                          'test_date': 'Test Date', 'result': 'Test Result', 'email_type': 'Email Type',
                                          'test_type': 'Test Type',
                                          'email_received_in_week_num': 'Calendar Week Email Was Received In',
                                          'test_date_occured_in_week_num': 'Calendar Week Test Occurred In'},
                                 inplace=True)

        logging.debug("Renamed email_contents_df columns")
        list_of_dataframes.append(email_contents_df)

    if len(list_of_dataframes) > 1:
        master_dataframe = concat(list_of_dataframes)

    if len(list_of_dataframes) == 1:
        master_dataframe = list_of_dataframes[0]

    else:
        logging.warning("No Emails Could Be Processed")
        return None

    log_stats(master_dataframe)
    master_dataframe.to_csv("dumped_results.csv", index=False, encoding='utf-8', date_format="%Y/%m/%d")


if __name__ == "__main__":
    file_handler = logging.FileHandler('CovidProcessor_Execution_Logs.log', mode='w')
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s.%(msecs)03d %(levelname)s : %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            file_handler,
            console_handler
        ]
    )

    run_script(dev=False)
    input("Programme Finshed - Press enter to close")
