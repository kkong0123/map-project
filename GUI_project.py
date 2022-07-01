import json
import csv
from selenium import webdriver
import time
import folium
from folium import CustomIcon, plugins
import pandas as pd
import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from pathlib import Path
import webbrowser
#UI파일 연결
base_dir = os.path.dirname(os.path.abspath(__file__))
form_class = uic.loadUiType(base_dir + "//untitled.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        #버튼에 기능을 연결하는 코드
        self.pushButton.clicked.connect(self.button1Function)
        self.pushButton.clicked.connect(self.print_first)
        self.pushButton_2.clicked.connect(self.main)
        self.pushButton_3.clicked.connect(self.button3Function)
        self.pushButton_4.clicked.connect(self.button4Function)
        self.pushButton_4.clicked.connect(self.print_first)
        self.pushButton_5.clicked.connect(self.server)

        self.setWindowTitle('지도변환 v1.0.1-release')

    def main(self):

        # 이름, 주소, 전화번호 등이 있는 행 찾기 
        for idx, val in enumerate(data): 
            if '이름' in val:
                meta_idx = idx
                break

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

        if '호차' in data[meta_idx]:
            num_idx = data[meta_idx].index("호차")

        if '부모님 전화번호' in data[meta_idx]:
            parent_num_idx = data[meta_idx].index("부모님 전화번호")

        name = []
        age = []
        gender = []
        adress = []
        phone = []
        num = []
        parent_phone = []
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
        try:    
            for i in range(meta_idx + 1, len(data)):
                num.append(data[i][num_idx])
        except Exception:
            for i in range(meta_idx + 1, len(data)):
                num.append(' ')     
            pass
        try:    
            for i in range(meta_idx + 1, len(data)):
                parent_phone.append(data[i][parent_num_idx])
        except Exception:
            for i in range(meta_idx + 1, len(data)):
                parent_phone.append(' ')     
            pass
        
        print(num)
        # 옵션 생성
        options = webdriver.ChromeOptions()
        # 창 숨기는 옵션 추가
        options.add_argument("headless")
        if getattr(sys, 'frozen', False):
            chromedriver_path = os.path.join(sys._MEIPASS, "./chromedriver")
            driver = webdriver.Chrome(chromedriver_path, options=options)
        else:
            driver = webdriver.Chrome('./chromedriver',options=options)
            
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
            print("{0}/{1} {2}".format(i+1, len(adress), location))

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
                    zoom_start=17)
        plugins.Fullscreen(
        position='topright',
        title='Expand me',
        title_cancel='Exit me',
        force_separate_button=True
        ).add_to(m)
        plugins.LocateControl().add_to(m)

        error_list = []
        for i in range(len(adress)):
            for j in range(1,26):
                if num[i] ==  str(j) + '호차':
                    image_icon = ".//image//{}.png".format(j)
                    icon1 = CustomIcon(
                        image_icon,
                        icon_size=(35, 35)
                    )
                    break
                else:
                    icon1 =folium.Icon('red', icon='star')

            if first_location[i].isalpha() or second_location[i].isalpha() == True:
                error_list.append(name[i])
                
            else:
                popup = str(name[i])+'<br/>' + str(age[i]) + str(gender[i]) + '<br/>' + str(adress[i]) + '<br/>' + str(phone[i]) + '<br/>' + "부모님: " + str(parent_phone[i])
                iframe = folium.IFrame(popup, width = 150, height=160)
                popup = folium.Popup(iframe)
                    

                folium.Marker([first_location[i], second_location[i]], popup, icon = icon1, tooltip = name[i]).add_to(m)



        print("전체 인원: {}\n누락된 인원:{}".format(len(adress), len(error_list)))
        print("[누락] 다음 인원의 주소를 다시 확인해주세요 : {}".format(error_list))
        self.textBrowser.append("\n--지도 추출 완료--")
        self.textBrowser.append("\n전체 인원: {}\n누락된 인원:{}".format(len(adress), len(error_list)))
        self.textBrowser.append("\n[누락] 다음 인원의 주소를 다시 확인해주세요 : {}".format(error_list))
        m
        m.save(save_path + '//{0}.html'.format(file_name))

    def button1Function(self) : # 엑셀로 열 때 함수
        print("엑셀 파일 가져옴")
        global fname
        global current_path
        global file_name
        global data
        fname = QFileDialog.getOpenFileName(self)
        current_path = fname[0].replace('/','//')
        file_name = Path(current_path).stem
        # excel 의 값을 list 로 변환
        df = pd.read_excel(current_path) 
        data = df.values.tolist() # 엑셀 내용 (1번째 행은 안 가져옴)
        data.insert(0, df.columns.tolist()) # 엑셀의 필드 값(1번째 줄) 가져와서 data 리스트의 0번째 인덱스에 추가 
        print(data)

    def button4Function(self): # csv로 열 때 함수
        global fname
        global file_name
        global data
        fname = QFileDialog.getOpenFileName(self)
        current_path = fname[0].replace('/','//')
        file_name = Path(current_path).stem
        print("CSV 파일 가져옴")
        data = list()
        f = open(current_path,'r',encoding='cp949')
        rea = csv.reader(f)
        for row in rea:
            data.append(row)
        f.close
        print(data)

    def button3Function(self):  
        global save_path
        save = QFileDialog.getExistingDirectory(self)
        save_path = save.replace('/','//')
        self.textBrowser.append("저장위치\n{}".format(save))

    def print_first(self):
        self.textBrowser.setPlainText("엑셀파일을 가져왔습니다.\n {}\n".format(fname[0]))
        print("start")

    def server(self):
        webbrowser.open("https://app.netlify.com/login/email")
        # 크롤링 막아둔듯 
        
if __name__ == "__main__" :        
    app = QApplication(sys.argv)
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
    