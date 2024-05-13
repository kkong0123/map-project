# 엑셀 데이터를 활용한 지도 시각화 프로그램

**Skills**

<img src="https://img.shields.io/badge/python-3776AB?style=for-the-badge&logo=HTML5&logoColor=white"> <img src="https://img.shields.io/badge/pyqt-41CD52?style=for-the-badge&logo=pyqt&logoColor=white"> <img src="https://img.shields.io/badge/selenium-43B02A?style=for-the-badge&logo=selenium&logoColor=white"> <img src="https://img.shields.io/badge/folium-77B829?style=for-the-badge&logo=folium&logoColor=white"> <img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white">

본 프로그램은 엑셀 또는 csv 파일에 담겨있는 정보를 가져와 주소를 추출하여 지도에 정보를 표시해 주는 프로그램입니다.

## 프로젝트 개요

학원 버스를 운영하시는 분께서 학생들의 픽업 위치를 한눈에 볼 수 있는 프로그램이 있으면 좋겠다고 말씀을 하셨습니다. 그래서 엑셀 파일 내 주소 정보를 활용하여 지도에 마커를 표시하는 프로그램을 개발해 봤습니다.

## 엑셀 파일 초기 설정 방법

프로그램 실행 전 몇 가지 초기설정이 필요합니다. 실제 학생 리스트 파일을 사용하기에는 개인정보 이슈가 있으므로 [서울시 강서구 공공주택 현황]이라는 공공데이터 엑셀 파일을 예시로 가져와 설명하겠습니다.  


1. 우선 사용할 파일을 열어 필드에서 **'주소'**, **'이름'**, '나이', '성별', '전화번호', '부모님 전화번호', '호차'에 해당하는 값을 찾습니다. 

   아래 예시를 보면 '단지명'은 '이름'에 해당하고 '도로명주소'는 '주소'에 해당, '연 락 처'는 '전화번호'라는 값에 해당합니다. 나머지 필드(나이, 성별, 부모님 전화번호, 호차)에는 매칭되는 데이터가 없기 때문에 생략해도 괜찮습니다.
   - 필수 항목은 **'이름'**,  **'주소'** 이므로 나머지 값(성별, 나이 등)은 없어도 무관합니다.
   

<img width="80%" alt="스크린샷 2022-05-12 오후 2 19 05" src="https://user-images.githubusercontent.com/104900966/167999305-e19c876b-d6a9-405f-9e9b-456f3c6aee5a.png">


2. 이 값들을 띄어쓰기하지 않고 다음과 같이 수정합니다.
<img width="80%" alt="스크린샷 2022-05-12 오후 2 19 39" src="https://user-images.githubusercontent.com/104900966/167999408-6a57d7b2-4c1c-436e-803e-b68d63776580.png">

3. 수정한 파일을 저장하여 사용합니다.

이런 과정을 거치는 이유는 본 프로그램이 ['주소', '이름', '나이', '성별', '전화번호', '부모님 전화번호', '호차']라는 단어가 있는 열을 찾아 작동하는 방식이기 때문입니다.

GUI 사용법
--

1. 오픈할 파일의 형태에 맞게 엑셀 파일 가져오기 or csv 파일 가져오기 버튼을 누른 후

2. 저장위치 설정 버튼을 눌러 저장 경로를 설정

3. 시작 버튼을 누른 후 잠시 대기, 데이터를 읽고 지도를 만드는 과정이므로 시간이 걸릴 수 있습니다. 
 
<img width="50%" alt="스크린샷 2022-05-12 오후 2 45 40" src="https://user-images.githubusercontent.com/104900966/168000387-f7d3b60e-6a17-47ad-915c-2351119deaec.png">


시연 영상
==
아래 시연 영상에서는 임의로 만든 학생명단 파일을 예시로 하였습니다.

<img width="80%" alt="스크린샷 2024-05-13 오후 3 31 23" src="https://github.com/kkong0123/map-project/assets/104900966/8c1108d9-40be-4920-9d08-81b39c8de77f">


https://github.com/kkong0123/map-project/assets/104900966/c12c285c-6141-46cb-a522-98b2a886b1c6

결과물
==
- 파일명.html 파일로 지도가 저장되며, 버스 호차(1호차, 2호차..)에 맞게끔 아이콘으로 마커가 표시됩니다.
- 아이콘을 클릭하면 세부정부가 나타납니다.

<img width="100%" alt="" src="https://github.com/kkong0123/map-project/assets/104900966/08ee64ec-7c7b-4c0d-b8c1-3aae12342162">

- 누락된 인원은 엑셀 파일 내 주소 값이 잘못 입력된 경우(오타)이며, 다시 확인 해달라는 메시지를 출력하게 했습니다.

<img width="50%" alt="스크린샷 2024-05-08 오후 9 35 25" src="https://github.com/kkong0123/map-project/assets/104900966/0a88e46c-dcea-46e9-9a2e-806a8f4d0ce6">

