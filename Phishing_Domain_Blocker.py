# Title: Automated Domain Blocker
# Creator: Hemal Maniar
# Date: 11/2/2022

# Libraries for Selenium
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.chrome.options import Options as CO
from selenium.webdriver.chrome.service import Service as CS
import time
from getpass import getpass

# Libraries for Outlook 
import win32com.client
import os
from datetime import datetime, timedelta
import eml_parser
import json
import datetime as dt

# Colour
import termcolor
os.system("color")

PATH = "C:\\Program Files (x86)\\"

yWarn = termcolor.colored("[WARNING]", "yellow")
bInfo = termcolor.colored("[INFO]", "blue")
gSuccess =  termcolor.colored("[SUCCESS]", "green")
rError = termcolor.colored("[ERROR]", "red")

def domainFetcher():
    day = datetime.today()
    start_time = day.replace(hour=0, minute=0, second=0).strftime("%Y-%m-%d %H:%M %p")
    end_time = day.replace(hour=23, minute=59, second=59).strftime("%Y-%m-%d %H:%M %p")
    
    # Outlook
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except:
        outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNameSpace("MAPI")

    alerts = mapi.Folders("Phish Alert").Folders("Inbox").Items
    alerts = alerts.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")

    PATH = os.getcwd()
    count = 1
    for message in list(alerts):
        try:
            attachment = message.Attachments.Item(1)
            attachment.SaveASFile(os.path.join(PATH, attachment.FileName))
            with open(attachment.FileName, "rb") as phish:
                rawmail = phish.read()

            ep = eml_parser.EmlParser()
            email = ep.decode_email_bytes(rawmail)
            parsed_email = json.dumps(email, indent=4, sort_keys=True, default=str)
            json_email = json.loads(parsed_email)

            # Email Header information
            emailFrom = str(json_email["header"]["from"]).split("@")
            emailFrom = emailFrom[1]

            domainCheck = domainChecker(emailFrom)
            
            # Write to file
            if domainCheck == None:
                print(f"{yWarn} {emailFrom} Cannot be blocked.")
            
            elif domainCheck == False:
                with open("Domain_List.txt", "a") as domainList:
                    domainList.write(f"{emailFrom}\n")
                    print(f"{gSuccess} [{count}] {emailFrom}")
                    count += 1
                    domainList.close()
            
            elif domainCheck == True:
                print(f"{yWarn} {emailFrom} Domain already exists.")
                continue
            
            # Delete attachment
            if os.path.exists(attachment.FileName):
                os.remove(attachment.FileName)
            
            else:
                print("")
        except:
            continue

def domainChecker(emailFrom):
    with open("Domain_List.txt", "r") as reader:
        list = reader.readlines()
        if (emailFrom+"\n") in list:
            reader.close()
            return True
        else:
            reader.close()
            return False

def domainListBlocker(emID, pwID, browser, duoAuth):
    if browser == "1":
        option = CO()
        option.add_argument("--log-level=3")
        option.add_argument("window-size=1920,1080")
        if duoAuth == "3":
            option.headless = False
        else:
            option.headless = True
        serv = CS(PATH + "chromedriver.exe")
        driver = wd.Chrome(service=serv, options=option)
    elif browser == "2":
        option = Options()
        option.headless = True
        option.add_argument("window-size=1920,1080")
        if duoAuth == "3":
            option.headless = False
        else:
            option.headless = True
        serv = Service(PATH + "geckodriver.exe")
        driver = wd.Firefox(options=option, service=serv)
            
    # Begin driver
    driver.get("https://admin.microsoft.com")
    print(f"\n{bInfo}\t\t Running browser.")

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "i0116")))
    email = driver.find_element(By.ID, "i0116")
    email.send_keys(emID)
    email.send_keys(Keys.RETURN)
    print(f"{bInfo}\t\t Entered email.")

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "passwordInput")))
    password = driver.find_element(By. ID, "passwordInput")
    password.send_keys(pwID)
    password.send_keys(Keys.RETURN)
    print(f"{bInfo}\t\t Entered password.")
    
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "duo_iframe")))
    duoFrame = driver.find_element(By. XPATH, "//iframe[@id='duo_iframe']")
    driver.switch_to.frame(duoFrame)

    # Duo authentication method
    if duoAuth == "1":    
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//button[contains(., 'Send Me a Push')]")))
        driver.find_element(By.XPATH, "//button[contains(., 'Send Me a Push')]").click()
        driver.switch_to.default_content()
        print(f"[STATUS]\t Waiting for Duo authentication...")
    elif duoAuth == "2":
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//button[contains(., 'Call Me')]")))
        driver.find_element(By.XPATH, "//button[contains(., 'Call Me')]").click()
        driver.switch_to.default_content()
        print(f"[STATUS]\t Waiting for Duo authentication...")
    elif duoAuth == "3":
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//button[contains(., 'Enter a Passcode)]")))
        driver.find_element(By.XPATH, "//button[contains(., 'Enter a Passcode')]").click()
        driver.switch_to.default_content()
        print(f"[STATUS]\t Waiting for Duo authentication...")
    
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
    print(f"{gSuccess}\t 2FA authenticated.")
    print(f"{bInfo}\t\t Logged in successfully.")    

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//a[@name='Exchange']")))
    driver.find_element(By. XPATH, "//a[@name='Exchange']").click()
    driver.switch_to.window(driver.window_handles[1])

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "Mailflow_id")))
    print(f"{bInfo}\t\t Opened Microsoft Exchange admin center.")
    driver.find_element(By.ID, "Mailflow_id").click()
    driver.find_element(By.ID, "transportrules").click()

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ecp-frame")))
    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@id='ecp-frame']"))
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//table/tbody/tr[14]/td[2]")))
    rule = driver.find_element(By.XPATH, "//table/tbody/tr[14]/td[2]")
    action = ActionChains(driver)
    action.double_click(rule)
    action.perform()    
    driver.switch_to.window(driver.window_handles[2])
    driver.maximize_window()

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "ResultPanePlaceHolder_contentContainer")))
    ruleEditor = driver.find_element(By.XPATH, "//table[@class='RuleParametersTable'][1]/tbody/tr[7]/td[3]")
    ruleEditor.click()

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "dlgModalError")))
    domainField = driver.find_element(By.ID, "ResultPanePlaceHolder_contentContainer_TransportRuleEditor_ctl03_ctl00_listview_InputBar_TextBox")
    domainList = open("Domain_List.txt", "r")
    count = 1
    while True:
        domain = domainList.readline()
        domain = domain.split("\n")
        domainField.send_keys(domain[0])
        domainField.send_keys(Keys.RETURN)
        if not domain[0]:
            break       
        print(f"{gSuccess}\t [{count}] {domain[0]} Successfully added to Domain Blocklist.")
        count += 1
    domainList.close()
    
    with open("Domain_List.txt", "r+") as resetDomainList:
        resetDomainList.truncate(0)
        resetDomainList.close()
    print(f"{bInfo}\t\t Domain List reset.")

    driver.find_element(By.ID, "dlgModalError_OK").click()
    driver.find_element(By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCommit").click()
    time.sleep(10)
    print(f"{gSuccess}\t Domain blocklist updated successully.")

    driver.switch_to.window(driver.window_handles[1])
    driver.find_element(By.ID, "O365_MainLink_Me").click()
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//a[@id='mectrl_body_signOut']")))
    driver.find_element(By.XPATH, "//a[@id='mectrl_body_signOut']").click()
    print(f"{bInfo}\t\t Logged out successfully.")

    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "login_workload_logo_text")))
    driver.quit()

def selector():
    print("\nAutomated Phishing Domain Blocker for O365 Exchange Center")
    choose = str(input("\n1. Fetch phishing domains from Outlook.\n2. Run domain list blocker\n\nYour Response: "))
    if choose == "1":
        domainFetcher()
        choose2 = str(input("\nDo you want to run domain list blocker? Y or N: "))
        choose2 = choose2.lower()
        if choose2 == "y":
            login()
        elif choose2 == "n":
            pexit()
        else:
            print("{rError}\t\t Invalid choice. Choose the correct option and try again.\n")
            selector()
    elif choose == "2":
        login()
    else:
        print("{rError}\t\t Invalid choice. Choose the correct option and try again.\n")
        selector()

def pexit():
    print("\nThank you!")
    exit(0)

def login():
    emID = str(input("Email: "))
    pwID = getpass("Password: ")
    browser = str(input("\nChoose your desired browser\n1. Chrome\n2. Firefox\n\nYour Response: "))
    if browser != "1" and browser != "2":
        print(f"{rError}\t\t Invalid choice. Choose the correct option and try again.\n")
        login()
    else:
        duoMethod(emID, pwID, browser)

def duoMethod(emID, pwID, browser):
    duoAuth = str(input("\nChoose your Duo authentication method\n1. Send Me a Push.\n2. Call Me.\n3. Enter a Passcode.\n\nYour Response: "))
    if duoAuth == "1" or duoAuth == "2" or duoAuth == "3":
        domainListBlocker(emID, pwID, browser, duoAuth)
    else:
        print(f"{rError}\t\t Invalid choice. Choose the correct option and try again.\n")
        duoMethod(emID, pwID, browser)

if __name__ == "__main__":
    selector()