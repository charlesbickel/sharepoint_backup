import os
import subprocess
import logging
from time import sleep


def run(full_loc, new_name, url):
    logger = logging.getLogger("sp.imagecapture")

    capture = r'c:\sharepoint_backup\resources\phantomjs\bin\phantomjs.exe ' \
              r'c:\sharepoint_backup\resources\phantomjs\bin\rasterize.js'

    filename = full_loc + f'\\{new_name}.pdf'
    options = '10in*11in'
    cmd = f'{capture} "{url}" "{filename}" "{options}"'
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)

    logger.info(f"Taking screenshot of {new_name} at {url}")
    attempts = 0
    while attempts < 10:
        p_status = p.wait()
        if p_status == 1:
            logger.warning(f"URL Not Found... Retrying {new_name} in 60 seconds.")
            attempts += 1
            sleep(60)
        else:
            logger.info(f"{new_name} Screenshot Completed...")
            break


if __name__ == '__main__':
    run(r"\\NetworkFolderLocation\SP_Archive\Intangibles\TestFolder",
        'TestFolder',
        'URLtoSpecificForm')
