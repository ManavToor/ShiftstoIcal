#!/usr/bin/env python

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
import json
import datetime
from dateutil.relativedelta import relativedelta
import smtplib
import ssl
from dateutil import parser
from icalendar import Calendar, Event
import dropbox

error_code = ''
teams = (('TEAM_9900213d-659d-4eba-ae87-d3e0889dcfe7', 'Greenbriar'),
         ('TEAM_7d78e379-c45b-4036-9a97-43b4c9a38878', 'Terry Miller'),
         ('TEAM_55ab2d30-2473-49bd-88e8-f64677a122e6', 'Greenbriar'),
         ('TEAM_da932026-6b67-415f-8388-e21f00ef67bf', 'Earnscliffe'),
         ('TEAM_8ea0e85d-6f67-4b19-88fb-7d513a522e75', 'Greenbriar'),
         ('TEAM_9b0a8719-638e-423f-b5a1-33d22d8fd055', 'Greenbriar'),
         ('TEAM_9f54d297-dcb8-4082-bd92-fb08fef7b2c7', 'Terry Miller'),
         ('TEAM_5244d34c-4da9-4299-8144-1e2610d71ffe', 'Gore Meadows'),
         ('TEAM_28816b16-e3a7-424a-b68f-bf19930b52a5', 'Greenbriar'),
         ('TEAM_6d9643f1-348f-43b4-b8e7-112e5ab982ad', 'CLTC'))


def get_shifts():
    desired_capabilities = DesiredCapabilities.CHROME
    desired_capabilities["goog:loggingPrefs"] = {"performance": "ALL"}

    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--allow-running-insecure-content')

    user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/5' \
                 '37.36'
    chrome_options.add_argument(f'user-agent={user_agent}')
    driver = webdriver.Chrome(options=chrome_options, desired_capabilities=desired_capabilities)

    url = 'https://teams.microsoft.com/'
    driver.get(url)

    time.sleep(5)
    ms_login = driver.find_element(By.ID, 'i0116')
    ms_login.send_keys('MY EMPLOYEE EMAIL')

    time.sleep(1)
    next_btn = driver.find_element(By.ID, 'idSIButton9')
    next_btn.click()

    time.sleep(4)
    password = driver.find_element(By.ID, 'passwordInput')
    password.send_keys('PASSWORD')


    time.sleep(1)
    sign_in_btn = driver.find_element(By.ID, 'submitButton')
    sign_in_btn.click()

    time.sleep(3)
    do_not_stay_signed_in_btn = driver.find_element(By.ID, 'idBtn_Back')
    do_not_stay_signed_in_btn.click()

    driver.implicitly_wait(50)
    driver.switch_to.frame(driver.find_element(By.NAME, 'embedded-page-container'))
    arenas_list = driver.find_element(By.CLASS_NAME, 'ms-List-page')
    arenas = arenas_list.find_elements(By.CLASS_NAME, 'ms-List-cell')
    # arena_name = arenas[0].find_element(By.CLASS_NAME, '_listItemName--3p1XV')
    # print(arena_name.text)
    view_shifts = arenas[0].find_element(By.TAG_NAME, 'button')
    view_shifts.click()
    driver.switch_to.default_content()

    time.sleep(10)
    driver.switch_to.frame(driver.find_element(By.NAME, 'embedded-page-container'))
    calender_btn = driver.find_element(By.XPATH, "//button[@title='Week']")
    calender_btn.click()
    time.sleep(1)
    select_month_btn = driver.find_element(By.NAME, 'Month')
    select_month_btn.click()

    def get_shifts(month_indicator='this_month.json', dump=False):
        today = datetime.datetime.today()
        month = today.strftime('%m')

        if month_indicator == 'next_month.json':
            month = today + relativedelta(months=+1)
            month = month.strftime('%m')
        elif month_indicator == 'last_month.json':
            month = today + relativedelta(months=-1)
            month = month.strftime('%m')

        logs = driver.get_log("performance")

        for log in logs:
            network_log = json.loads(log["message"])["message"]
            if "Network.responseReceived" in network_log["method"] and "params" in network_log.keys():
                try:
                    if "json" in network_log["params"]["response"]["mimeType"]:
                        for team in teams:
                            if network_log['params']['response']['url'].startswith(
                                    f'https://flw.teams.microsoft.com/svc-nam1/api/users/me/dataindaterange?teamIds='
                                    f'{team[0]}&startTime=2023-' + month + '-01T'):
                                request_id = network_log['params']['requestId']
                                body = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
                                body = json.loads(body['body'])
                                body = body['shifts']
                                print(body)

                                shifts = {'shifts': []}
                                if dump:
                                    with open(month_indicator, 'w') as outfile:
                                        for shift in body:
                                            shifts['shifts'].append(shift)
                                        print(shifts)
                                        json.dump(shifts, outfile)
                                else:
                                    with open(month_indicator, 'r') as outfile:
                                        prev_shifts = json.load(outfile)
                                        for shift in body:
                                            prev_shifts['shifts'].append(shift)
                                    with open(month_indicator, 'w') as outfile:
                                        print(prev_shifts)
                                        json.dump(prev_shifts, outfile)
                                break
                except KeyError:
                    pass

    def next_month(dump=False):
        time.sleep(3)
        if dump:
            get_shifts(dump=True)
        else:
            get_shifts()

        time.sleep(7)
        next_month_btn = driver.find_element(By.XPATH, "//button[@title='Go to next month']")
        next_month_btn.click()

        time.sleep(2)
        if dump:
            get_shifts('next_month.json', True)
        else:
            get_shifts('next_month.json')

        time.sleep(2)
        last_month_btn = driver.find_element(By.XPATH, "//button[@title='Go to previous month']")
        last_month_btn.click()
        time.sleep(2)
        last_month_btn = driver.find_element(By.XPATH, "//button[@title='Go to previous month']")
        last_month_btn.click()

        time.sleep(2)
        if dump:
            get_shifts('last_month.json', True)
        else:
            get_shifts('last_month.json')

        time.sleep(2)
        next_month_btn = driver.find_element(By.XPATH, "//button[@title='Go to next month']")
        next_month_btn.click()

        time.sleep(5)

    for i in range(len(arenas)):
        if i == 0:
            next_month(True)
        elif i == len(arenas) - 1:
            next_month()
            break
        else:
            next_month()

        time.sleep(5)
        arena_chooser_icon = driver.find_element(By.ID, 'team_Picker_Icon_Id')
        arena_chooser_icon.click()

        time.sleep(3)
        teams_list = driver.find_elements(By.CLASS_NAME, '_listItem--3HUhk')
        teams_list[i + 1].click()

        time.sleep(10)

    driver.switch_to.default_content()
    driver.quit()


def to_ical():
    cal = Calendar()

    def get_data(shift_file):
        with open(shift_file, 'r') as input_file:
            schedule = json.load(input_file)

        for i in schedule['shifts']:
            event = Event()

            event.add('summary', i['notes'])
            event.add('dtstart', parser.parse(i['startTime']))
            event.add('dtend', parser.parse(i['endTime']))
            for j in teams:
                if j[0] == i['teamId']:
                    event.add('location', j[1])
                    print(j[1])
                    break
            cal.add_component(event)

    get_data('last_month.json')
    get_data('this_month.json')
    get_data('next_month.json')

    with open('schedule.ics', 'wb') as cal_file:
        cal_file.write(cal.to_ical())


def upload_to_dropbox():
    dbx = dropbox.Dropbox('TjXdRB4FtVwAAAAAAAAAAQD9aEi7XNgHvLZ-_9jSatLs2F_XeNP6qCm_eDfb-D6c')

    dbx.files_upload(open('schedule.ics', 'rb').read(), '/schedule.ics', mode=dropbox.files.WriteMode.overwrite)


def exceptions(error, function):
    exception_time = datetime.datetime.now()
    smtp_server = "smtp.gmail.com"
    sender_email = "SENDER EMAIL"
    receiver_email = "MY EMAIL"
    password = "PASSWORD"
    message = f"""\
    Subject: Raspberry Pi Exception

    The following exception:
    {error}
    has occurred on the raspberry pi at:
    {exception_time}
    in the function:
    {function}
    """

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, 587) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)


if __name__ == '__main__':
    try:
        get_shifts()
        to_ical()
        upload_to_dropbox()
    except Exception as e:
        error_code = e
        exceptions(e, 'get_shifts()')
        quit()
