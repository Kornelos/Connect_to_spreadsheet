from time import sleep
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.firefox.options import Options as FirefoxOptions
import pandas as pd
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import os.path
import gspread
from sys import argv
from secrets import garmin_email, garmin_password, spreadsheet_url, FTP


login_url = "https://sso.garmin.com/sso/signin?service=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&webhost=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&source=https%3A%2F%2Fconnect.garmin.com%2Fsignin%2F&redirectAfterAccountLoginUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&redirectAfterAccountCreationUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&gauthHost=https%3A%2F%2Fsso.garmin.com%2Fsso&locale=en_US&id=gauth-widget&cssUrl=https%3A%2F%2Fstatic.garmincdn.com%2Fcom.garmin.connect%2Fui%2Fcss%2Fgauth-custom-v1.2-min.css&privacyStatementUrl=https%3A%2F%2Fwww.garmin.com%2Fen-US%2Fprivacy%2Fconnect%2F&clientId=GarminConnect&rememberMeShown=true&rememberMeChecked=false&createAccountShown=true&openCreateAccount=false&displayNameShown=false&consumeServiceTicket=false&initialFocus=true&embedWidget=false&generateExtraServiceTicket=true&generateTwoExtraServiceTickets=false&generateNoServiceTicket=false&globalOptInShown=true&globalOptInChecked=false&mobile=false&connectLegalTerms=true&showTermsOfUse=false&showPrivacyPolicy=false&showConnectLegalAge=false&locationPromptShown=true&showPassword=true#"
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


def download_csv():
    """Downloads a csv file from garmin connect profile specified in secrets.py"""
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.download.dir", os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")
    options = FirefoxOptions()
    options.add_argument("--headless") # in background mode
    with webdriver.Firefox(firefox_profile=fp, options=options) as driver:
        wait = WebDriverWait(driver, 10)
        driver.get(login_url)
        driver.find_element_by_name("username").send_keys(garmin_email)
        driver.find_element_by_name("password").send_keys(garmin_password + Keys.RETURN)
        sleep(4)
        driver.get("https://connect.garmin.com/modern/activities?activityType=cycling")
        sleep(1)
        wait.until(presence_of_element_located((By.CSS_SELECTOR, ".export-btn")))
        driver.find_element_by_css_selector(".export-btn").click()
        sleep(1)


def main():
    iterations = None
    if argv.count is 2:
        assert argv[1].isnumeric() is False, "USAGE: number of iterations or nothing for full list iteration"
        iterations = int(argv[1])
    download_csv()
    df = pd.read_csv("Activities.csv")[['Activity Type', 'Date', 'Time', 'Distance', 'Training Stress Score®',
                                        'Power', 'Normalized Power® (NP®)', 'Avg HR', 'Max HR',
                                        ]]
    df['Date'] = df['Date'].str[:10]
    df = df.sort_values('Distance', ascending=False).drop_duplicates('Date').sort_index()  # longest workout per day
    os.remove('Activities.csv')
    # logging into Google Sheet
    credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_url(spreadsheet_url)
    worksheet = sh.get_worksheet(0) # open latest created worksheet
    # worksheet population with data
    i = 0
    if iterations is None:
        iterations = len(df.index)
    for index, row in df.iterrows():
        if i is iterations:
            break
        date = datetime.strptime(row['Date'], '%Y-%m-%d')
        sheet_format_date = date.strftime('%#d.%#m.%Y')
        try:
            cell = worksheet.find(sheet_format_date)
            worksheet.update_cell(cell.row, 16, row['Time'])
            worksheet.update_cell(cell.row, 17, row['Distance'])
            worksheet.update_cell(cell.row, 19, row['Training Stress Score®'])
            worksheet.update_cell(cell.row, 20, round(row['Normalized Power® (NP®)'] / FTP, 3))
            worksheet.update_cell(cell.row, 21, row['Power'])
            worksheet.update_cell(cell.row, 22, row['Normalized Power® (NP®)'])
            worksheet.update_cell(cell.row, 23, row['Avg HR'])
            worksheet.update_cell(cell.row, 24, row['Max HR'])
        except gspread.exceptions.GSpreadException:
            continue
        i += 1


if __name__ == '__main__':
    main()
