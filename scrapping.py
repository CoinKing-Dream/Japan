
# Python program to demonstrate
# selenium
dateJapan = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"]
dateShort = ['Jan', 'Feb', 'Mar', 'Apr', "May", 'Jun', 'Jul','Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# import webbrowser
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
import datetime
from selenium.webdriver.common.by import By
from configparser import ConfigParser
import calendar
import tkinter as tk
from tkinter import messagebox
import requests
import json

import gspread
from oauth2client.service_account import ServiceAccountCredentials



wait_time = 50
# import the parameters from external file(.ini)
config = ConfigParser()
config.read('config.ini')

# room id information
temp_roomID_Group = config.get('room_id', 'roomID_Group')
temp_roomID = config.get('room_id', 'roomID')

roomID_Group = temp_roomID_Group.split(',')
roomID = temp_roomID.split(',')

for i in range(len(roomID)): roomID[i] = roomID[i].strip()
for i in range(len(roomID_Group)): roomID_Group[i] = roomID_Group[i].strip()


# url setting
first_website_url = config.get('website_url', 'first_website_url')
token_website_url = config.get('website_url', 'token_website_url')
api_website_url = config.get('website_url', 'api_website_url')
second_website_url = config.get('website_url', 'second_website_url')
googlesheet_url = config.get('website_url', 'googlesheet_url')
api_key = config.get('website_url', 'api_key')

# username, password parameter
username_first_website = config.get('first_website', 'username')
password_first_website = config.get('first_website', 'password')

password_second_website = config.get('second_website', 'password')



# import autonmation time
automation_first_time = config.get('automation_times', 'automation_first_time')
automation_second_time = config.get('automation_times', 'automation_second_time')

interval = config.get('automation_times', 'interval_time')
interval = int(interval)





def get_refresh_token(invite_code):
    end_point = "/authentication/setup"
    url = api_website_url + end_point
    headers = {
        "accept": "application/json",
        "code": invite_code
    }
    response = requests.get(url=url, headers=headers)
    refresh_token = response.json()['refreshToken']
    
    return refresh_token



def get_token(refresh_token):
    end_point = "/authentication/token"
    url = api_website_url + end_point
    headers = {
        "accept": "application/json",
        "refreshToken": refresh_token
    }
    response = requests.get(url=url, headers=headers)
    token = response.json()['token']
    
    return token    



def get_result(token, params={}):
    end_point = "/inventory/rooms/availability"
    url = api_website_url + end_point
    
    header = {
        "accept": "application/json",
        "Token": token
    }
    response = requests.get(url=url, headers=header, params=params)
    result = response.json()
    
    return result

def write_data_into_googlesheet(room_number, room_data):
    
    try:
        # accessing Google Sheet

        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = ServiceAccountCredentials.from_json_keyfile_name('myapi.json', scopes)
        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_url(googlesheet_url)
        worksheet = spreadsheet.get_worksheet(0)
        print("Google Sheet accessed!")
        

    except Exception as e:
        print(f'Google Sheet Do Not access! {e}')    

    sleep(1)

    count = len(roomID_Group)
    worksheet.update_cell(1, 1, 'Date')
    sleep(0.5)

    for item in range(count):
        worksheet.update_cell(1, int(item + 2), str(roomID_Group[item]))
        sleep(0.5)
    print("-----------------------")
    
    date_array = list(room_data[0]['availability'].keys())
    value_array = []

    print("Sheet Data is Updating...., Wait...")
    for item in range(count):
        value_array.append(list(room_data[room_number[item]]['availability'].values()))
    

    for item in range(len(date_array)):
        worksheet.update_cell(int(item + 2), 1, str(date_array[item]))
        print(f'{item + 2} Row, {1} Column -> {str(date_array[item])} ')
        sleep(1)
        for i in range(count):
            is_empty = '0' if value_array[i][item] == False else '1'   
        
            try:
                worksheet.update_cell(int(item + 2), int(i + 2), is_empty)
                sleep(1)
                print(f'{item + 2} Row, {i + 2} Column -> {is_empty} ')
            except Exception as e:
                print(f'Row Update Error! {e}')
        
        
    for item in range(35):
        for i in range(count + 1):
        
            try:
                worksheet.update_cell(int(item + 2 + len(date_array)), int(i + 1), "")
                print(f'{item + 2 + len(date_array)} Row, {i + 1} Column -> Delete ')
                sleep(1)
            except Exception as e:
                print(f'Row Update Error! {e}')

    print("-----------------------")

    print(f'{room_number} Rooms Updated!')

        


def main():

    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"]) 
        options.add_argument('--headless=new')
        browser = webdriver.Chrome(options=options)

        # starting Browser
        
        browser.get(first_website_url)
        print(f'Browser Started!')
    except Exception as e:
        print(f'The Browser Do Not Start! {e}')    
    print("-----------------------")

    wait = WebDriverWait(browser, wait_time)

    userName = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="settingformid"]/div/div[3]/div[1]/div/div/div/div/input')))
    userName.clear()
    userName.send_keys(username_first_website)

    passWord = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="settingformid"]/div/div[3]/div[2]/div/div/div/div/input')))
    passWord.clear()
    passWord.send_keys(password_first_website)
    
    loginButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="settingformid"]/div/div[4]/button')))
    
    loginButton.click()
    print("Login Button Clicked!")
    print("-----------------------")
    browser.implicitly_wait(20)

    browser.get(token_website_url)

    try:
        deleteInviteButtons = wait.until(EC.visibility_of_all_elements_located((By.XPATH, f'//*[@id="invitetokenlisttable"]/tbody/*')))
    except Exception as e:
        print(f'{e}')
    sleep(1)

    count = 0
    count = len(deleteInviteButtons)
    if not (0 >= count and count < 100): count = 0

    print(count)
    
    for item in range(count):
        try:
            deleteInviteButton = (webdriver, 2).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="invitetokenlisttable"]/tbody/tr/td[6]')))
            browser.execute_script("arguments[0].scrollIntoView();", deleteInviteButton)
            deleteInviteButton.click()
            print(f'{item + 1} Invite Code Delete!')
        except Exception as e:
            pass
    sleep(1)        

    try:
        tokenButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="settingformid"]/div/div/div/div[2]/button')))
        browser.execute_script("arguments[0].scrollIntoView();", tokenButton)
        tokenButton.click()
        print("Token Button Clicked!")
    except Exception as e:
        print(f'Token Butoon Do Not Click! {e}')

    sleep(1)
    for item in range(7):
        try:
            checkButton = wait.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="scopeModal"]/div/div/div[2]/table/tbody/tr[{item + 1}]/td[4]/div/input')))
            browser.execute_script("arguments[0].scrollIntoView();", checkButton)
            if not checkButton.is_selected(): checkButton.click()
            print(f'{item + 1} Check Button Clicked!')
        except Exception as e:
            print(f'Check Butoon Do Not Click! {e}')
    

    try:
        scopeModalButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="scopeModalSubmit"]')))
        browser.execute_script("arguments[0].scrollIntoView();", scopeModalButton)
        scopeModalButton.click()
        print("scopeModal Button Clicked!")
    except Exception as e:
        print(f'scopeModal Butoon Do Not Click! {e}')

    
    try:
        invitation_code = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="invitetokenlisttable"]/tbody/tr/td[1]/span')))
        
        print("Invite Code Got!")
    except Exception as e:
        print(f'Invite Code Do Not Get! {e}')

    # browser.close()
    print('')
    print('--------------------------------------------------')
    print('')

    refresh_token = get_refresh_token(invitation_code.text)
    print ("Refresh Token Got!")

    token = get_token(refresh_token)
    print ("Token Got!")
    
    params = {
        "startDate": startDate,
        "endDate": endDate
    }

    result = get_result(token, params)
    # print (f"Result : {result}")

    # sleep(1000)
    # print(f'Result Success ----->  {result['success']}')
    
    if result['success']:
        count = len(roomID)

        room_number = []
        for i in range(count): room_number.append(0)

        for i in range(count):
            for j in range(count):
                if str(roomID_Group[i]) == str(result['data'][j]['propertyId']):
                    room_number[int(i)] = int(j)
                    break
                       
        write_data_into_googlesheet(room_number, result['data'])
                

    print("Google Sheet Data Insert End!")
    print("-------------------------------")


    #------------- ACO Website  ---------------------

    try:
        # accessing Google Sheet

        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = ServiceAccountCredentials.from_json_keyfile_name('myapi.json', scopes)
        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_url(googlesheet_url)
        worksheet = spreadsheet.get_worksheet(0)

        print("Google Sheet accessed!")

    except Exception as e:
        print(f'Google Sheet Do Not access! {e}')    

    sleep(1)


    print("-----------------------")

    try:
        data = worksheet.get_all_values()

        print("Data are gotten")
        print('-----------------------------')
        print('------------  ACO Website  -----------------')
    except Exception as e:
        print(f'Sheet Do Not access to read the data {e}')
    


    for roomNumber in range(len(roomID)):
        # loading duration
        # browser.delete_all_cookies()
        print(str(roomID[roomNumber]) + " Browser start ...")

        options = webdriver.ChromeOptions()
        options.add_experimental_option('detach', True)
        options.add_argument('--headless=new')

        browser = webdriver.Chrome(options=options)
        browser.get(second_website_url)
        wait = WebDriverWait(browser, wait_time)
        
        userName = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="id"]')))
        userName.clear()
        userName.send_keys(roomID[roomNumber])

        passWord = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pass"]')))
        passWord.clear()
        passWord.send_keys(password_second_website)

        try:
            loginButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login"]')))
            loginButton.click()
            print(f'Login Clicked!') 
        except Exception as e:
            print(f'Login Do Not Clicked {e}') 


        try:
            roomInfoButton = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div/a')))
            roomInfoButton.click()
            print(f'RoomInfoButton Clicked!') 
        except Exception as e:
            print(f'RoomInfoButton Do Not Clicked {e}') 

        
        oldCalendar = ""
        browser.implicitly_wait(10)
        browser.switch_to.window(browser.window_handles[-1])

        for i in range(len(data)):
            sleep(1)
            if i > 0: 
                
                data[i][0] = data[i][0].replace("年", "-")
                data[i][0] = data[i][0].replace("月", "-")
                data[i][0] = data[i][0].replace("日", "")

                date = data[i][0].split('-')
                
                try:
                    date[0] = date[0].strip()
                    date[1] = date[1].strip()
                    date[2] = date[2].strip()

                except Exception as e:
                    print(f'{e}')

                    if date[0] == "":
                        date[0] = "2023"
                    if date[1] == "":
                        date[1] = "12"
                    if date[2] == "":
                        date[2] = "01"


                if len(date[1]) == 1:
                    date[1] = '0{}'.format(date[1])
                
                if len(date[2]) == 1:
                    date[2] = '0{}'.format(date[2])

                newCalendar = '{}-{}'.format(date[0], date[1])

                
                

            
                if newCalendar != oldCalendar: 
                    oldCalendar = newCalendar
                    try:
                        compButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submit"]')))
                        browser.execute_script("arguments[0].scrollIntoView();", compButton)
                        sleep(1)
                        browser.execute_script("arguments[0].click();", compButton)
                        print("CompButton Clicked ------ A")

                    except ElementClickInterceptedException as e:
                        print(f'CompButton is not clickable at the moment. Waiting for it to become clickable... {e}') 

                    try:
                        returnButton = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/p/a')))
                        returnButton.click()
                        print(f'ReturnButton Clicked!') 
                    except Exception as e:
                        print(f'RoomInfoButton Do Not Clicked {e}') 


                    try:
                        calendarButton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@href="?hiduke={}"]'.format(newCalendar))))
                        print('//*[@href="?hiduke={}"]'.format(newCalendar))
                        # Scroll to the element (if needed)
                        browser.execute_script("arguments[0].scrollIntoView();", calendarButton)
                        browser.execute_script("arguments[0].click();", calendarButton)
                        print("CalendarButton Clicked") 
                       
                    except ElementClickInterceptedException as e:
                        print(f'Element is not clickable at the moment. Waiting for it to become clickable... {e}') 

                    sleep(1)

                try:    
                    if data[i][roomNumber + 1] == "0":
                        num = 3
                    elif data[i][roomNumber + 1] == "1":
                        num = 2
                    else:
                        num = 3
                    
                except Exception as e:
                    num = 3
                
                try:
                    waitID = '//*[@name="item[{}-{}-{}-{}][status]"]/option[{}]'.format(roomID[roomNumber], date[0], date[1], date[2], num)
                    optionButton = WebDriverWait(browser, wait_time).until(EC.visibility_of_element_located((By.XPATH, waitID)))
                    browser.execute_script("arguments[0].scrollIntoView();", optionButton)
                    optionButton.click()
                    print(f'{waitID} --> Select ---> {data[i][roomNumber + 1]}')
                except Exception as e:
                    print(f'Don not Select! {e}') 

                compDate = datetime.datetime(int(date[0]), int(date[1]), int(date[2]))

                print(f'{date[0]}-{date[1]}-{date[2]} --------- {endDate_eight_month}')
                if(f'{date[0]}-{date[1]}-{date[2]}' == endDate_eight_month): break
        
        try:
            compButton = WebDriverWait(browser, wait_time).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submit"]')))
            browser.execute_script("arguments[0].scrollIntoView();", compButton)
            sleep(1)
            browser.execute_script("arguments[0].click();", compButton)
            print("CompButton Clicked ------ B") 
        except ElementClickInterceptedException as e:
            print(f'CompButton is not clickable at the moment. Waiting for it to become clickable... {e}') 


        try:
            returnButton = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/p/a')))
            returnButton.click()
           
            print(f'ReturnButton Clicked!') 
        except Exception as e:
            print(f'RoomInfoButton Do Not Clicked {e}') 
            
        browser.quit()

        sleep(3)

    browser.quit()

    print('')
    print(' -------------------------------- ')
    print(f'Data Scrapping Work finished at {datetime.datetime.now()}')
    print(' -------------------------------- ')


while True:
    current_time = datetime.datetime.now()
    first_comp_time = automation_first_time.split('-')
    second_comp_time = automation_second_time.split('-')

    config.read('config.ini')

    automation_first_time = config.get('automation_times', 'automation_first_time')
    automation_second_time = config.get('automation_times', 'automation_second_time')
    interval = config.get('automation_times', 'interval_time')
    interval = int(interval)

    # start and end date
    print("")
    year = str(current_time.year)
    mon = str(current_time.month)
    day = str(current_time.day)

    if len(mon) == 1:mon = '0{}'.format(mon)    
    if len(day) == 1:day = '0{}'.format(day)

    startDate = '{}-{}-{}'.format(year, mon, day)
    print(f'Starting Date: ------>  {startDate}')


    a_year_from_now = datetime.datetime(current_time.year, current_time.month, int(day)) + datetime.timedelta(days=30 * 12)
    
    year = str(a_year_from_now.year)
    mon = str(a_year_from_now.month)
    day = str(calendar.monthrange(int(year), int(mon))[1])

    if len(mon) == 1:mon = '0{}'.format(mon)
    if len(day) == 1:day = '0{}'.format(day)

    endDate = '{}-{}-{}'.format(year, mon, day)
    print(f'Ending Date: ------>  {endDate}')


    eight_months_from_now = datetime.datetime(current_time.year, current_time.month, 1) + datetime.timedelta(days=30 * 8)
    
    year = str(eight_months_from_now.year)
    mon = str(eight_months_from_now.month)
    day = str(calendar.monthrange(int(year), int(mon))[1])

    if len(mon) == 1:mon = '0{}'.format(mon)
    if len(day) == 1:day = '0{}'.format(day)

    endDate_eight_month = '{}-{}-{}'.format(year, mon, day)
    print(f'Ending Date: ------>  {endDate_eight_month}')
    

    print("")
    print(f'Waiting ----->  Current Time: {current_time}')
   

    # endDate = "2023-11-28"
    # print(int(first_comp_time[0]))
    # print(int(current_time.hour))
    # print(int(first_comp_time[1]))
    # print(int(current_time.minute))

    if (int(first_comp_time[0]) == int(current_time.hour) and int(first_comp_time[1]) == int(current_time.minute)) or (int(second_comp_time[0]) == int(current_time.hour) and int(second_comp_time[1]) == int(current_time.minute)):    

        print("")
        print("")
        print("------------------------------------------------")

        print(f'Start Time: {current_time}')

        main()

        print('-------------- To see the updating time, open the Log.txt file  ----------')
        
        with open('Log.txt', 'r') as myfile:  
            result = myfile.read()

        with open('Log.txt', 'w') as myfile:  
            myfile.write(result + '\n\n' + f' Start Time ---> {current_time},  End Time   ---> {datetime.datetime.now()}')
        
        print("")
        print('---------------------- During Time -------------------------')
        print(f' Start Time ---> {current_time}')
        print(f' End Time   ---> {datetime.datetime.now()}')
        print("")
    
    sleep(interval)

    

# if startDate > endDate:
        #     # Create a tkinter window
        #     window = tk.Tk()

        #     # Hide the main window
        #     window.withdraw()

        #     # Show the alert dialog
        #     messagebox.showinfo("Alert", "Wrong Date!")

        #     # Close the tkinter window
        #     window.destroy()

        #     pass
        # else: