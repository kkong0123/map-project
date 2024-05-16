import os
import sys
import csv
import time
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from PyQt5.QtWidgets import *
from PyQt5 import uic
import folium
from folium import CustomIcon, plugins

# UI파일 연결
base_dir = os.path.dirname(os.path.abspath(__file__))
form_class = uic.loadUiType(base_dir + "/map.ui")[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 변수 초기화
        self.data = None
        self.current_path = None
        self.file_name = None
        self.save_path = None

        # 버튼을 함수에 연결
        self.pushButton.clicked.connect(self.button_excel)
        self.pushButton.clicked.connect(self.display_files)
        self.pushButton_2.clicked.connect(self.main)
        self.pushButton_3.clicked.connect(self.button_save)
        self.pushButton_4.clicked.connect(self.button_csv)
        self.pushButton_4.clicked.connect(self.display_files)
        self.setWindowTitle('지도 시각화 프로그램 v2.0')

    def main(self):
        # 데이터가 없는 경우 경고 메시지 출력
        if not self.data:
            self.textBrowser.append("데이터가 없습니다. 먼저 파일을 열어주세요.")
            return
        address, name, age, gender, phone, num, parent_phone = self.extract_data()  # 데이터 추출
        first_location, second_location = self.get_coordinates(address) # 위도, 경도 추출
        self.create_map(name, age, gender, address, phone, num, parent_phone, first_location, second_location) # 지도 생성

    def extract_data(self):
        # 데이터 추출 함수
        for idx, val in enumerate(self.data): # 데이터에서 이름, 주소, 전화번호 등이 있는 행을 찾아서 해당 행의 인덱스(meta_idx) 찾기
            if '이름' in val: 
                meta_idx = idx
                break

        # 엑셀 데이터 정보를 추출하여 각각의 리스트로 저장
        address = self.extract_column(meta_idx, self.extract_idx(meta_idx, "주소"))
        name = self.extract_column(meta_idx, self.extract_idx(meta_idx, "이름"))
        age = self.extract_column(meta_idx, self.extract_idx(meta_idx, "나이"))
        gender = self.extract_column(meta_idx, self.extract_idx(meta_idx, "성별"))
        phone = self.extract_column(meta_idx, self.extract_idx(meta_idx, "전화번호"))
        num = self.extract_column(meta_idx, self.extract_idx(meta_idx, "호차"))
        parent_phone = self.extract_column(meta_idx, self.extract_idx(meta_idx, "부모님 전화번호"))
    
        return address, name, age, gender, phone, num, parent_phone
    
    def extract_idx(self, meta_idx, string):
        # 주어진 문자열(string)이 데이터의 특정 행(meta_idx)에 있는 경우, 해당 문자열의 열 인덱스를 반환하는 함수
        if string in self.data[meta_idx]:
            return self.data[meta_idx].index(string) 
            
    def extract_column(self, meta_idx, col_idx):
        # 특정 열 데이터 추출하는 함수
        column_data = []
        for i in range(meta_idx + 1, len(self.data)):
            try:
                value = self.data[i][col_idx]
                column_data.append(value)
            except Exception:
                column_data.append('')
        return column_data

    def get_coordinates(self, address):
        # 위도, 경도 값 추출 함수
        options = webdriver.ChromeOptions()
        # options.add_argument("headless") # 크롬 창 백그라운드에서 동작하기 
        driver = webdriver.Chrome(options=options)

        url = 'https://address.dawul.co.kr/'
        driver.get(url)
        time.sleep(1)
        driver.switch_to.window(driver.window_handles[1])
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        first_location = [] # 위도 값 리스트
        second_location = [] # 경도 값 리스트

        for i in range(len(address)):
            driver.find_element(By.CSS_SELECTOR, "#input_juso").click()
            driver.find_element(By.CSS_SELECTOR, "#input_juso").send_keys(address[i])
            driver.find_element(By.CSS_SELECTOR, "#btnSch").click()
            time.sleep(1)

            location = driver.find_element(By.CSS_SELECTOR, "#insert_data_5").text
            location = location.split(',')
            print("{0}/{1} {2}".format(i+1, len(address), location))

            location_1 = location[1].split(':')[1]
            location_2  = location[0] .split(':')[1]

            first_location.append(location_1)
            second_location.append(location_2)

            print(location_1)
            print(location_2 )

        driver.quit()
        return first_location, second_location

    def create_map(self, name, age, gender, address, phone, num, parent_phone, first_location, second_location):
        # 지도 생성 함수
        m = folium.Map(location=[first_location[0], second_location[0]], zoom_start=17)
        plugins.Fullscreen(position='topright', title='Expand me', title_cancel='Exit me', force_separate_button=True).add_to(m)
        plugins.LocateControl().add_to(m)

        error_list = []
        for i in range(len(address)):
            icon1 = self.get_icon(num[i])
            if first_location[i].isalpha() or second_location[i].isalpha():
                error_list.append(name[i])
            else:
                popup_content = ""
                if name[i]:
                    popup_content += f"{name[i]}"
                    if gender[i]:
                        popup_content += f"({gender[i]})"
                    popup_content += "<br>"
                if age[i]:
                    popup_content += f"나이: {age[i]}<br>"
                if phone[i]:
                    popup_content += f"학생: {phone[i]}<br>"
                if parent_phone[i]:
                    popup_content += f"부모님: {parent_phone[i]}<br>"
                if address[i]:
                    popup_content += f"<b>{address[i]}</b>"

                iframe = folium.IFrame(popup_content, width=250, height=120)
                popup = folium.Popup(iframe)
                folium.Marker([first_location[i], second_location[i]], popup, icon=icon1, tooltip=name[i]).add_to(m)

        self.display_results(len(address), len(error_list), error_list)
        m.save(self.save_path + f'/{self.file_name}.html')

    def get_icon(self, num):
        # 호차에 따른 아이콘 설정 함수
        for j in range(1, 26):
            if num == f"{j}호차":
                return CustomIcon(f"./image/{j}.png", icon_size=(35, 35))
        return folium.Icon('red', icon='star')

    def display_results(self, total, errors, error_list):
        # 결과 출력하는 함수
        self.textBrowser.append("\n--지도 추출 완료--")
        self.textBrowser.append(f"\n전체 인원: {total}\n누락된 인원: {errors}")
        if error_list:
            self.textBrowser.append(f"\n[누락] 다음 인원의 주소를 다시 확인해주세요 : {error_list}")

    def display_files(self):
        # 파일 경로 출력 함수
        self.textBrowser.setPlainText(f"파일을 가져왔습니다.\n {self.current_path}\n")
        
    def button_excel(self):
        # 엑셀 파일 열기
        self.current_path, _ = QFileDialog.getOpenFileName(self)
        if not self.current_path:
            return
        self.file_name = Path(self.current_path).stem
        df = pd.read_excel(self.current_path)
        self.data = df.values.tolist()
        self.data.insert(0, df.columns.tolist())

    def button_csv(self):
        # csv 파일 열기
        self.current_path, _ = QFileDialog.getOpenFileName(self)
        if not self.current_path:
            return
        self.file_name = Path(self.current_path).stem
        with open(self.current_path, 'r', encoding='cp949') as f:
            rea = csv.reader(f)
            self.data = list(rea)

    def button_save(self):
        # 저장 경로 설정
        self.save_path = QFileDialog.getExistingDirectory(self)
        self.textBrowser.append(f"저장위치\n{self.save_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
