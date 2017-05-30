import xlrd
import os
import logging.config
import requests
import getpass
from requests_ntlm import HttpNtlmAuth
from lib import imgcapture
from lib import dl_attachments


# Logger info
logging.config.fileConfig('resources\\logging.conf')
logger = logging.getLogger('sp.expenses')

# Open and read ExpenseReports.xlsx. Read from 1st worksheet
workbook = xlrd.open_workbook("resources\\ExpenseReports.xlsx")
worksheet = workbook.sheet_by_index(0)
total_rows = worksheet.nrows

# Gathering NTLM Credentials to download PDF directly
username = input('Username: ')
password = getpass.getpass()

folder_location = r"\\NetworkFolderLocation\sp_archive" + "\\Expenses\\"


def save_pdf(url, full_loc, new_name):
    res = requests.get(url, auth=HttpNtlmAuth(username, password))
    with open(full_loc + '\\' + new_name + '.pdf', 'wb') as f:
        f.write(res.content)


def chk_makedir(id_excel, url, name):
    new_name = id_excel + "_" + name.rstrip()
    full_loc = folder_location + new_name
    if not os.path.exists(full_loc):
        logger.info(f"Making {new_name} directory at {full_loc}")
        os.makedirs(full_loc)
    if len(os.listdir(full_loc)) < 2:
        if url.endswith('.pdf'):
            logger.info("URL is a PDF? Just downloading it...")
            save_pdf(url, full_loc, new_name)
        else:
            imgcapture.run(full_loc, new_name, url)
            dl_attachments.get(full_loc, url)


# Add %20 for every space to make a functional url
def convert_url(xml):
    ending = xml.replace(" ", "%20")
    url = "http://sharepointwebsite.com/forms/Expenses/" + ending
    return url


# Remove .xml from name
def slice_name(self):
    name = self[:-4]
    return name


# Read 1st two columns of ExpenseReports.xlsx
def main():
    logger.info("---- Program Started ----")

    for i in range(1, total_rows):
        id_excel = (int(worksheet.cell(i, 0).value))
        url = convert_url(worksheet.cell(i, 1).value)
        name = slice_name(worksheet.cell(i, 1).value)
        chk_makedir(str(id_excel), url, name)

    logger.info("---- Program Finished ----")


if __name__ == '__main__':
    main()
