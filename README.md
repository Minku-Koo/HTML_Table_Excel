# HTML_Table_Excel
## Extract HTML Table and Input a Table Data to Excel



📚 This Library is applied HTML_Table_Extractor


📌 Library Name : Table_Excel


📌 Created Date : 27/Aug/2020


📌 Updated Date : 11/Mar/2021


📌 Author : Minku Koo


📌 E-Mail : corleone@kakao.com


📌 Version : 1.1.4


📌 Keywords : 'Excel', 'Table', 'HTML', 'Crawling', 'Selenium', 'Extractor'


----------------------------------------------------------------------------------------------------------------------------

# ⚙ How to Use?
 ```
from HTML_Table_Excel import Table_Excel

# ENG
TableExcel = Table_Excel( URL_list <type=(String)list>, ChromeDriver Path <type=String>)
TableExcel.makeExel_abs( Excel File Path <type=String>, Table Header Color by Hex <type=String> (Default=F8E0EC) )
TableExcel.makeExel_sep( Excel File Path <type=String> )

# KOR
TableExcel = Table_Excel( URL <리스트>, 크롬 드라이버 경로 <문자열>)
TableExcel.makeExel_abs( 엑셀 파일 경로 <문자열>, 테이블 헤더 색깔 - 16진수 <문자열> (Default=F8E0EC) )
TableExcel.makeExel_sep( 엑셀 파일 경로 <문자열> )
```


# 📝 Explains

- HTML table 태그의 데이터를 수집 및 변형하여 Excel 파일로 만들어주는 라이브러리 입니다.
- 엑셀 파일에는 링크, 페이지 제목이 포함되어 있습니다.
- 해당 웹 페이지의 모든 테이블을 수직으로 정렬시켜 표시합니다.
- 각 테이블의 헤더는 색을 달리하여 표시해줍니다.

✔ makeExel_sep() 함수는 테이블을 그대로 보여줍니다. rowspan, colspan에서 병합이 이루어지지 않습니다.

✔ makeExel_abs() 함수는 테이블의 병합을 그대로 구현합니다. rowspan, colspan의 병합이 엑셀에서도 동일하게 이루어집니다.

✔ 중첩 테이블, 가로 정렬 테이블도 모두 표시해줍니다.


# 📢 You should check this

- You should check your ChromeDriver version

- Also, You have to check, that your Chrome Browser Version and your ChromeDriver version is same


----------------------------------------------------------------------------------------------------------------------------

# 💡 Here is Examples

## 📍 Sample 1 (What is different between makeExel_sep() and makeExel_abs()?) 
*(URL : https://www.weather.go.kr/weather/observation/currentweather.jsp)*


### 🖥 Web Page 
![weather-web2](https://user-images.githubusercontent.com/25974226/110779854-60cb0800-82a7-11eb-8571-faefdcbe8316.PNG)

### 🔍 Table_Excel -> makeExel_sep() 
![seq-weather2](https://user-images.githubusercontent.com/25974226/110779541-ffa33480-82a6-11eb-9fce-b2911b81a371.PNG)

### 🔍 Table_Excel -> makeExel_abs() 
![abs-weather2](https://user-images.githubusercontent.com/25974226/110779532-fd40da80-82a6-11eb-9ebc-8ad580711d0d.PNG)


## 📍 Sample 2 (How about Table in table or horizontal arangement tables?) 
*(URL : http://www.kweather.co.kr/kma/kma_digital.html)*


### 🖥 Web Page 
![weather-web](https://user-images.githubusercontent.com/25974226/110779543-ffa33480-82a6-11eb-9f6e-8bc7a1c1d682.PNG)

### 🔍 Table_Excel -> makeExel_abs() 
![abs-weather](https://user-images.githubusercontent.com/25974226/110779549-00d46180-82a7-11eb-9e4e-3d1365a4a725.PNG)


## 📍 Sample 3 (Table in table case) 
*(path : ./sample_html/innerTable_Sample.html)*


### 🖥 HTML 
![inner-html](https://user-images.githubusercontent.com/25974226/110779538-ff0a9e00-82a6-11eb-9853-1df9b37610fb.PNG)

### 🔍 Table_Excel -> makeExel_abs() 

![abs-html](https://user-images.githubusercontent.com/25974226/110779544-003bcb00-82a7-11eb-95ea-e3c128921fb0.PNG)



## 📍 Sample 4 (Horizontal arangement tables case) 
*(path : ./sample_html/horizontal_table_sample.html)*


### 🖥 HTML 

![horizon-html](https://user-images.githubusercontent.com/25974226/110779536-fe720780-82a6-11eb-8865-a09c57adc256.PNG)


### 🔍 Table_Excel -> makeExel_abs() 

![abs-html2](https://user-images.githubusercontent.com/25974226/110779545-003bcb00-82a7-11eb-8f1a-59d1b30ea7b1.PNG)


