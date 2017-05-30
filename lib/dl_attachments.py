import re
import glob
import shutil
import logging
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from time import sleep


def get(full_loc, url):

    logger = logging.getLogger("sp.dl_attachments")

    # Set all Chromedriver options
    chromeoptions = webdriver.ChromeOptions()
    chromeoptions.add_argument("--safebrowsing-disable-download-protection")
    chromeoptions.add_experimental_option("prefs", {
        "download.default_directory": full_loc,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    logger.info("Starting Selenium...")
    chromedriver = r"C:\sharepoint_backup\resources\chromedriver.exe"
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeoptions)
    driver.set_window_size(1120, 550)
    driver.get(url)
    html_source = driver.page_source

    attempts = 0
    while attempts < 6:
        if attempts == 5:
            logger.critical(f"Could not access this site {url} "
                            f"location {full_loc} ")
            quit()
        try:
            if driver.find_element_by_xpath("//*[contains(text(),'This site can’t be reached')]"):
                logger.warning('Page did not load. Waiting 5 minutes and trying again...')
                attempts += 1
                sleep(300)
                driver.get(url)
        except NoSuchElementException:
            break

    '''
    Find the attachments by xpath
    Select xpath for class with a specific class name, but does not contain a specific title
    Double click to to download and accept the alert message.
    '''
    logger.info(f"Downloading attachments for {full_loc} ...")

    try:
        m = re.search('(name=\"FormControl_\w{2}_\w{2}_\w{3}_\w{2}_\w{4}_\w{4}_\w{2}\")\s(class=\"(\w{2}_\w{16}_\w\s\w{2}_\w{16}_\w)\")', html_source)
        for content in driver.find_elements_by_xpath(f'//*[contains(@class, "{m.group(3)}") '
                                                     'and not(contains(@title, "Click here to attach a file"))]'):
            try:
                actionchains = ActionChains(driver)
                actionchains.double_click(content).perform()
                Alert(driver).accept()
                sleep(.8)
            except NoAlertPresentException:
                logger.warning(f"Could not find alert window for {full_loc} {url}")
                pass
    except AttributeError:
        logger.info("No attachment to download.")

    # Look for "unfinished" download files. Sleep if exist.
    sleep(2)
    while glob.glob(full_loc + '\\*.crdownload' or full_loc + '\\*.tmp'):
        sleep(3)

    driver.quit()
    clean_appdata()


# Chromedriver issue: Builds up files in appdata and never picks up after itself... what a slob.
def clean_appdata():
    cleandir = r"C:\Users\charles.bickel\AppData\Local\Temp\scoped_dir*"
    for folder in glob.iglob(cleandir, recursive=False):
        shutil.rmtree(folder, ignore_errors=False)


if __name__ == '__main__':
    get(r"\\NetworkFolderLocation\SP_Archive\Intangibles\TestFolder",
        'URLtoSpecificForm')
