# Automatic Phishing Domain Blocker

An automatic phishing domain blocker that is designed to fetch email domain from your "Phish Alert" mailbox right into your Microsoft Office 365 Exchange Admin Center. 

## Requirements

1. Python3 <= 3.9 - Doesn't work with latest version 3.10 as eml_parser Python library is not supported.
2. Pip 
3. Place "chromedriver.exe" and "geckodriver.exe" in your desired path. _Defualt set to - "C:\Program Files (x86)\"_.

## Dependencies

1. pip install selenium - Selenium webdriver automation
2. pip install eml_parser - Parsing .eml files
3. pip install pypiwin32 - Interact with Outlook application
