# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:light
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.5'
#       jupytext_version: 1.13.8
#   kernelspec:
#     display_name: python-venv
#     language: python
#     name: python-venv
# ---
from dataclasses import replace
from tkinter import Frame
import requests
import json
import csv
from selenium import webdriver
import time
import folium
import pandas as pd
import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

def main():
    # excel 의 값을 list 로 변환
    df = pd.read_excel(current_path) # images 폴더 위치 반환
    data = df.values.tolist()

    # +
    # csv로 열 때
    # data = list()
    # f = open("//Users//kkong0123//Desktop//python_program//project//이름.csv",'r',encoding='cp949')
    # rea = csv.reader(f)
    # for row in rea:
    #     data.append(row)
    # f.close
    # print(data)
    # -

    # 주소 이름 ..이 있는 행을 찾기 (모든 엑셀 파일의 똑같은 행 위치에 주소, 이름 .. 이 위치하는 것이 아니기에 -> 유동적으로 )
    for idx, val in enumerate(data): 
        if '이름' in val:
            meta_idx = idx
            break
    print(meta_idx)

    print(data)

    # +
    if '이름' in data[meta_idx]:
        name_idx = data[meta_idx].index("이름") 
        
    if '나이' in data[meta_idx]:
        age_idx = data[meta_idx].index("나이") 

    if '성별' in data[meta_idx]:
        gender_idx = data[meta_idx].index("성별") 

    if '주소' in data[meta_idx]:
        adress_idx = data[meta_idx].index("주소") # 주소 위치 인덱스를 찾아서 인덱스 저장

    if '전화번호' in data[meta_idx]:
        phone_idx = data[meta_idx].index("전화번호") 

    name = []
    age = []
    gender = []
    adress = []
    phone = []
    # 이름","나이","성별","주소","전화번호"

    try:
        for i in range(meta_idx + 1, len(data)):
            name.append(data[i][name_idx]) 
    except Exception:
        for i in range(meta_idx + 1, len(data)):
            name.append(' ')     
        pass

    try:
        for i in range(meta_idx + 1, len(data)):
            age.append(data[i][age_idx]) 
    except Exception:
        for i in range(meta_idx + 1, len(data)):
            age.append(' ')     
        pass

    try:    
        for i in range(meta_idx + 1, len(data)):
            gender.append(data[i][gender_idx]) 
    except Exception:
        for i in range(meta_idx + 1, len(data)):
            gender.append(' ')     
        pass
        
    try:
        for i in range(meta_idx + 1, len(data)):
            phone.append(data[i][phone_idx]) 
    except Exception:
        for i in range(meta_idx + 1, len(data)):
            phone.append(' ') 
        pass

    try:    
        for i in range(meta_idx + 1, len(data)):
            adress.append(data[i][adress_idx]) # adress 변수에 주소 저장
    except Exception:
        for i in range(meta_idx + 1, len(data)):
            adress.append(' ')     
        pass

    print(phone)

    # 옵션 생성
    options = webdriver.ChromeOptions()
    # 창 숨기는 옵션 추가
    options.add_argument("headless")

    driver = webdriver.Chrome('.//chromedriver',options=options)
    url = 'https://address.dawul.co.kr/'
    driver.get(url)
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    first_location = [] # 위도 값 리스트
    second_location = [] # 경도 값 리스트

    for i in range(len(adress)):
        driver.find_element_by_css_selector("#input_juso").click()
        driver.find_element_by_css_selector("#input_juso").send_keys(adress[i])
        driver.find_element_by_css_selector("#btnSch").click()
        time.sleep(1)

        location = driver.find_element_by_css_selector("#insert_data_5").text
        location = location.split(',')
        print(location)
        global location_1
        global location_2

        location_1 = location[1]
        location_2 = location[0]

        location_1 = location_1.split(':')[1]
        location_2  = location_2 .split(':')[1]

        first_location.append(location_1)
        second_location.append(location_2)

        print(location_1)
        print(location_2 )

    driver.quit()
    # driver.close()

    m = folium.Map(location=[first_location[0], second_location[0]],
                zoom_start=17, 
                width=750, 
                height=500)


    # +
    # # 서울 행정구역 json raw파일(githubcontent)
    # r = requests.get('https://raw.githubusercontent.com/southkorea/seoul-maps/master/kostat/2013/json/seoul_municipalities_geo_simple.json')
    # c = r.content
    # seoul_geo = json.loads(c)

    # m = folium.Map(
    #     location=[37.559819, 126.963895],
    #     zoom_start=11, 
    # )

    # folium.GeoJson(
    #     seoul_geo,
    #     name='지역구'
    # ).add_to(m)


    # +
    error_list = []

    for i in range(len(adress)):
        if first_location[i].isalpha() or second_location[i].isalpha() == True:
            error_list.append(name[i])
            
        else:
            popup = str(name[i])+'<br/>' + str(age[i]) + str(gender[i]) + '<br/>' + str(adress[i]) + '<br/>' + str(phone[i])
            iframe = folium.IFrame(popup, width = 150, height=160)
            popup = folium.Popup(iframe)
            folium.Marker([first_location[i], second_location[i]], popup, tooltip = name[i]).add_to(m)

        
    print("전체 인원: {}\n누락된 인원:{}".format(len(adress), len(error_list)))
    print("[누락] 다음 인원의 주소를 다시 확인해주세요 : {}".format(error_list))

    m

    # -

    m.save('TEST_MAP.html')


#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("untitled.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        #버튼에 기능을 연결하는 코드
        self.pushButton.clicked.connect(self.button1Function)
        self.pushButton_2.clicked.connect(self.button2Function)



    def button1Function(self) :
        print("엑셀 파일 가져오기")
        global fname
        global current_path
        fname = QFileDialog.getOpenFileName(self)
        current_path = fname[0].replace('/','//')

    def button2Function(self):
        self.textBrowser.setPlainText("dd")
        print("start")
        # xlsx으로 열 때
        # Excel 파일 불러오기
        main()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()

    
'''


# +
# xlsx으로 열 때
# Excel 파일 불러오기
current_path = os.path.dirname('__file__') # 현재 파일의 위치 반환
df = pd.read_excel(os.path.join(current_path, "2호차.xlsx")) # images 폴더 위치 반환

# excel 의 값을 list 로 변환

data = df.values.tolist()

# +
# csv로 열 때
# data = list()
# f = open("//Users//kkong0123//Desktop//python_program//project//이름.csv",'r',encoding='cp949')
# rea = csv.reader(f)
# for row in rea:
#     data.append(row)
# f.close
# print(data)
# -

# 주소 이름 ..이 있는 행을 찾기 (모든 엑셀 파일의 똑같은 행 위치에 주소, 이름 .. 이 위치하는 것이 아니기에 -> 유동적으로 )
for idx, val in enumerate(data): 
    if '이름' in val:
        meta_idx = idx
        break
print(meta_idx)

print(data)

# +
if '이름' in data[meta_idx]:
    name_idx = data[meta_idx].index("이름") 
    
if '나이' in data[meta_idx]:
    age_idx = data[meta_idx].index("나이") 

if '성별' in data[meta_idx]:
    gender_idx = data[meta_idx].index("성별") 

if '주소' in data[meta_idx]:
    adress_idx = data[meta_idx].index("주소") # 주소 위치 인덱스를 찾아서 인덱스 저장

if '전화번호' in data[meta_idx]:
    phone_idx = data[meta_idx].index("전화번호") 

name = []
age = []
gender = []
adress = []
phone = []
# 이름","나이","성별","주소","전화번호"

try:
    for i in range(meta_idx + 1, len(data)):
        name.append(data[i][name_idx]) 
except Exception:
    for i in range(meta_idx + 1, len(data)):
        name.append(' ')     
    pass

try:
    for i in range(meta_idx + 1, len(data)):
        age.append(data[i][age_idx]) 
except Exception:
    for i in range(meta_idx + 1, len(data)):
        age.append(' ')     
    pass

try:    
    for i in range(meta_idx + 1, len(data)):
        gender.append(data[i][gender_idx]) 
except Exception:
    for i in range(meta_idx + 1, len(data)):
        gender.append(' ')     
    pass
    
try:
    for i in range(meta_idx + 1, len(data)):
        phone.append(data[i][phone_idx]) 
except Exception:
    for i in range(meta_idx + 1, len(data)):
        phone.append(' ') 
    pass

try:    
    for i in range(meta_idx + 1, len(data)):
        adress.append(data[i][adress_idx]) # adress 변수에 주소 저장
except Exception:
    for i in range(meta_idx + 1, len(data)):
        adress.append(' ')     
    pass

print(phone)

# +
driver = webdriver.Chrome('//Users//kkong0123//Desktop//python_program//project//chromedriver')
url = 'https://address.dawul.co.kr/'
driver.get(url)
time.sleep(1)
driver.switch_to.window(driver.window_handles[1])
driver.close()
driver.switch_to.window(driver.window_handles[0])

first_location = [] # 위도 값 리스트
second_location = [] # 경도 값 리스트

for i in range(len(adress)):
    driver.find_element_by_css_selector("#input_juso").click()
    driver.find_element_by_css_selector("#input_juso").send_keys(adress[i])
    driver.find_element_by_css_selector("#btnSch").click()
    time.sleep(1)

    location = driver.find_element_by_css_selector("#insert_data_5").text
    location = location.split(',')
    print(location)

    location_1 = location[1]
    location_2 = location[0]

    location_1 = location_1.split(':')[1]
    location_2  = location_2 .split(':')[1]

    first_location.append(location_1)
    second_location.append(location_2)

    print(location_1)
    print(location_2 )

driver.close()
# -

m = folium.Map(location=[first_location[0], second_location[0]],
               zoom_start=17, 
               width=750, 
               height=500)


# +
# # 서울 행정구역 json raw파일(githubcontent)
# r = requests.get('https://raw.githubusercontent.com/southkorea/seoul-maps/master/kostat/2013/json/seoul_municipalities_geo_simple.json')
# c = r.content
# seoul_geo = json.loads(c)

# m = folium.Map(
#     location=[37.559819, 126.963895],
#     zoom_start=11, 
# )

# folium.GeoJson(
#     seoul_geo,
#     name='지역구'
# ).add_to(m)


# +
error_list = []

for i in range(len(adress)):
    if first_location[i].isalpha() or second_location[i].isalpha() == True:
        error_list.append(name[i])
        
    else:
        popup = str(name[i])+'<br/>' + str(age[i]) + str(gender[i]) + '<br/>' + str(adress[i]) + '<br/>' + str(phone[i])
        iframe = folium.IFrame(popup, width = 150, height=160)
        popup = folium.Popup(iframe)
        folium.Marker([first_location[i], second_location[i]], popup, tooltip = name[i]).add_to(m)

    
print("전체 인원: {}\n누락된 인원:{}".format(len(adress), len(error_list)))
print("[누락] 다음 인원의 주소를 다시 확인해주세요 : {}".format(error_list))

m

# -

m.save('TEST_MAP.html')

'''


