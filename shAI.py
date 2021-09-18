from selenium import webdriver
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import smtplib
import pyperclip
import pyautogui
import time

PATH = "D:\malek\chrome_driver\chromedriver.exe"
driver = webdriver.Chrome(PATH)

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("shAI.json", scope)
client = gspread.authorize(creds)

def click(xpath):
    while True:
        try:
            button = driver.find_element_by_xpath(xpath)
            button.click()
            time.sleep(1)
            break

        except:
            time.sleep(0.01)
            continue


def write(xpath, word):
    while True:
        try:
            search_box = driver.find_element_by_xpath(xpath)
            search_box.send_keys(word)
            time.sleep(1)
            break
        except:
            time.sleep(0.01)
            continue


def open_url(url):
    driver.get(url)

########################################
message = ["نرحب فيك من شاي للذكاء الاصطناعي 🤩","شكرا لتسحيلك في البرنامج الكامل من شاي تفاصيل البرنامج:","""برنامج يبدأ من الصفر وتعلم أساسيات البرمجة باستخدام لغة البايثون دون الحاجة لأي خلفية برمجية مسبقة 🚫
وصولاً لمواضيع متقدمة جداً مثل الرؤية الحاسوبية ومعالجة النصوص الطبيعية مروراً بجميع التفاصيل المهمة خلال هذه الرحلة.🧐
البرنامج يتكون من 5 مراحل تدريبية بما يزيد عن 150 ساعة تدريبية بطابع عملي من خلال تطبيق مشاريع حقيقية.💪
مراحل البرنامج:
= Coding with Python
= Data Science
= Hands-on Machine Learning
= Deep Dive into Deep Learning
= Advanced Topics: Natural Language Processing/ Computer Vision"""]
########################################
sheet = client.open(" التسجيل في برنامج الذكاء الاصطناعي المقدم من شاي - الدفعة الثانية (september program )").sheet1  # sheet name

empty_cell = sheet.cell(129, 20).value
sta_row = 40

while True:
    done = sheet.cell(sta_row, 17).value
    if done == empty_cell:
        way_of_contact = sheet.cell(sta_row, 11).value.split(', ')
        name=sheet.cell(sta_row, 2).value

        if 'WhatsApp' in way_of_contact:
            country_id = sheet.cell(sta_row, 7).value.split('+')[1]
            if sheet.cell(sta_row, 8).value[0] != '0':
                pho_num = sheet.cell(sta_row, 8).value
            else:
                pho_num = sheet.cell(sta_row, 8).value[1:]

            open_url(f'https://web.whatsapp.com/send?phone={country_id}{pho_num}')
            for i in message:
                pyperclip.copy(i)
                click('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
                pyautogui.hotkey("ctrl", "v")
                click('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]/button/span')

            sheet.update_cell(sta_row, 17, "whatsapp")
            sta_row = sta_row + 1
            time.sleep(1)

        elif 'Facebook' in way_of_contact:
            my_user_name = 'ai.juclub@gmail.com'  # facebook username
            password = 'JUai123321**'  # facebook pass

            mes = sheet.cell(sta_row, 10).value
            if '?id=' in mes:
                user_name = mes.split('=')[1]
            else:
                user_name = mes.split('.com/')[1]

            open_url('https://www.messenger.com/login.php?next=https%3A%2F%2Fwww.messenger.com%2Ft%2F100004007154698')
            write('//*[@id="email"]', my_user_name)
            write('//*[@id="pass"]', password)
            click('//*[@id="loginbutton"]')

            open_url(f'https://www.messenger.com/t/{user_name}')

            for i in message:

                pyperclip.copy(i)
                click("//div[starts-with(@id,'mount_0_0_')]/div/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/form/div/div[3]/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div/div/div")
                pyautogui.hotkey("ctrl", "v")
                click("//div[starts-with(@id,'mount_0_0_')]/div/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/form/div/div[3]/span[2]/div")

            sheet.update_cell(sta_row, 17, "messenger")
            sta_row = sta_row + 1
            time.sleep(1)


        elif 'Email' in way_of_contact:
            country_id = sheet.cell(sta_row, 7).value.split('+')[1]
            if sheet.cell(sta_row, 8).value[0] != '0':
                pho_num = sheet.cell(sta_row, 8).value
            else:
                pho_num = sheet.cell(sta_row, 8).value[1:]

            open_url(f'https://web.whatsapp.com/send?phone={country_id}{pho_num}')
            for i in message:
                pyperclip.copy(i)
                click('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
                pyautogui.hotkey("ctrl", "v")
                click('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]/button/span')

            sheet.update_cell(sta_row, 17, "whatsapp")
            sta_row = sta_row + 1
            time.sleep(1)


        else:
            sheet.update_cell(sta_row, 17, "do it handley")
            sta_row = sta_row + 1
            time.sleep(1)

    else:
        if done == 'donee':
            sheet.update_cell(sta_row, 17, "done")
            sta_row = sta_row + 1
            time.sleep(1)
        else:
            sta_row = sta_row + 1
            time.sleep(1)