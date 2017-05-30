# sharepoint_backup

This is a project I created to backup sharepoint libraries. The library information is exported into an excel document and then read by the Python program. The program uses phantomJS to take a PDF screenshot of each sharepoint form. Then it uses Selenium to go to the form and then downloads the attachments on the page. Python's requests library was not useable for this project since the attachment links were obfuscated by Javascript.
