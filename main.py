from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from json import loads,dumps
import time
from openpyxl import Workbook
import os
import getpass
import threading

url = "https://meet.google.com/dke-mjah-gnw"

class Functions():
    def __init__(self) -> None:
        while True:
            answer = input('請選擇功能:\n(1 離開')
            if answer == '1':
                pass
    def quit(self):
        roll_call_list.save_file()
        driver.quit()
        quit()

class Member_events():
    members = []
    def __init__(self)->None:
        while True:
            self.detection()
    def detection(self)->None:
        elements_for_now = driver.find_elements(By.CLASS_NAME,"zWGUib")
        members_for_now=[]
        for member in elements_for_now:
            try:
                members_for_now.append(member.text)
            except:
                pass
        if members_for_now != self.members:
            if len(members_for_now) > len(self.members):
                self.member_joined(members_for_now)
            else:
                self.member_left(members_for_now)
            self.members = members_for_now
            return
    def member_joined(self,members_for_now)->None:
        for member in members_for_now:
            if not member in self.members:
                roll_call_list.join(member)
                print(member+'加入了!')
                return
    def member_left(self,members_for_now)->None:
        for member in self.members:
            if not member in members_for_now:
                print(member+'離開了!')
                roll_call_list.left(member)
                return

class Roll_call_list(Workbook):
    __file_name = ''
    def __init__(self):
        self.__file_name = input("點名單檔案名稱:")+'.xlsx'
        self.create()
        self.save_file()
    def create(self):
        super().__init__()
        worksheet = self.active
        worksheet.merge_cells('A1:B1')
        worksheet['A1']='加入'
        worksheet.merge_cells('C1:D1')
        worksheet['C1']='離開'
        worksheet['E1']='點名紀錄'
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 20
        worksheet.column_dimensions['C'].width = 25
        worksheet.column_dimensions['D'].width = 20
    def save_file(self):
        while True:
            try:
                self.save(self.__file_name)
            except:
                input('無法儲存檔案，請關閉檔案後按enter重試')
            else:
                print('檔案儲存成功!')
                return
    def join(self,user_name:str):
        current_cell = 2
        worksheet = self.active
        while True:
            if worksheet['A'+str(current_cell)].value:
                current_cell+=1
                break
            worksheet['A'+str(current_cell)] = time.ctime()
            worksheet['B'+str(current_cell)] = user_name
            self.save_file()
            return
    def left(self,user_name:str):
        current_cell = 2
        worksheet = self.active
        while True:
            if worksheet['C'+str(current_cell)].value:
                current_cell+=1
                continue
            worksheet['C'+str(current_cell)] = time.ctime()
            worksheet['D'+str(current_cell)] = user_name
            self.save_file()
            return

class Login_Google():
    def __init__(self) -> None:
        driver.get(url)
        if os.path.isfile('config'):
            #add cookie
            try:
                self.login_with_cookie()
            except:
                print('登入憑證無效，重新建立中...')
                self.manual_login()
        else:
            print('未偵測到登入憑證，請手動登入')
            self.manual_login()
        driver.implicitly_wait(10)
        driver.get(url)
    def login_with_cookie(self):
        driver.delete_all_cookies()
        with open("config","r") as file:
            cookie = loads(file.read())
            for name,value in cookie.items():
                driver.add_cookie({"name":name,"value":value})
    def manual_login(self):
        driver.get("https://accounts.google.com/v3/signin/identifier?dsh=S-905868804%3A1676301694738053&_ga=2.207397329.35016544.1676301693-561019711.1676301693&continue=https%3A%2F%2Fmeet.google.com%3Fhs%3D193&ltmpl=meet&o_ref=https%3A%2F%2Fmeet.google.com%2F&flowName=GlifWebSignIn&flowEntry=ServiceLogin&ifkv=AWnogHdWVd-qRecSxEULl_n6vHYB07-UsP6Iqy7uBGNwMSXZ3GGXmR3rUzCtiuwX1w4tF4PMLDRIDQ")

        driver.implicitly_wait(30)

        # Enter your username and password
        username = driver.find_element(By.NAME,"identifier")
        username.send_keys(input('輸入您的gmail帳號:'))

        # Click the next button
        next_button = driver.find_element(By.ID,"identifierNext")
        next_button.click()

        # Wait for the password field to load
        # ...

        driver.implicitly_wait(10)

        passwd = driver.find_element(By.NAME,"Passwd")
        passwd.send_keys(getpass.getpass('輸入您的密碼:'))

        # Click the next button
        next_button = driver.find_element(By.ID,"passwordNext")
        next_button.click()
        driver.get(url)
        driver.implicitly_wait(10)
        web_cookies = driver.get_cookies()
        cookies = {}
        for cookie in web_cookies:
            cookies[cookie['name']]=cookie['value']
        with open('config','x')as file:
            file.write(dumps(cookies))
        self.login_with_cookie()


#preparation
print('初始化軟體...')
print('建立點名單...')
roll_call_list = Roll_call_list()
print('準備加入會議...')
#options.add_argument('--headless')
#options.add_argument('--disable-gpu')
driver=webdriver.Edge()
driver.maximize_window()

print('嘗試登入...')
while True:
    Login_Google()

    driver.implicitly_wait(4)
    try:
        driver.find_element(By.CLASS_NAME,"AjXHhf").click()
    except:
        print('登入失敗')
        os.remove('config')
        continue
    else: 
        link = driver.find_element(By.CLASS_NAME,"QJgqC")
        link.click()
        driver.implicitly_wait(4)
        driver.find_elements(By.CLASS_NAME,"boDUxc")[1].click()
        driver.implicitly_wait(4)
        break
print('會議加入成功!')
time.sleep(1)
os.system('cls')

threading.Thread(Member_events().__init__()).start()
threading.Thread(Functions().__init__()).start()
    
#driver.close()