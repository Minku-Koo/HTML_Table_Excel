# HTML_Table_Excel
Scrapping HTML Table and Input a Table Data to Excel

- This Library apply HTML_Table_Extractor

- Library Name : Table_Excel
- Created Date : 27/Aug/2020
- Updated Date : 07/Sep/2020
- Author : Minku Koo
- E-Mail : corleone@kakao.com
- Version : 1.1.3
- Keywords : 'Excel', 'Table', 'HTML', 'Crawling', 'Selenium', 'Extractor'



# How to Use?
 ```
TableExcel = Table_Excel( URL_list <type=(String)list>, ChromeDriver Path <type=String>)
TableExcel.makeExel_abs( Excel File Path <type=String>, Table Header Color by Hex <type=String> (Default=F8E0EC) )
TableExcel.makeExel_sep( Excel File Path <type=String> )
```





 * Please, Import these Library : HTML_Table_Extractor, BeautifulSoup4, openpyxl, time, selenium
 * You should check your ChromeDriver exist
 * Also, You have to check, that your Chrome Version and your ChromeDriver version is same

----------------------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------------------


- HTML의 table 태그의 데이터를 수집 및 변형하여 Excel 파일로 만들어주는 라이브러리 입니다.
- 엑셀 파일에는 링크, 페이지 제목이 포함되어 있습니다.
- 해당 웹 페이지의 모든 테이블을 수직으로 정렬시켜 표시합니다.
- 각 테이블의 헤더는 색을 달리하여 표시해줍니다.




# 사용법
 ```
TableExcel = Table_Excel( URL <리스트>, 크롬 드라이버 경로 <문자열>)
TableExcel.makeExel_abs( 엑셀 파일 경로 <문자열>, 테이블 헤더 색깔 - 16진수 <문자열> (Default=F8E0EC) )
TableExcel.makeExel_sep( 엑셀 파일 경로 <문자열> )
```



----------------------------------------------------------------------------------------------------------------------------
## 예시 Example 1 (Simple, Version 1.0.0)


- Absorption Excel View

<Example 1>

@web page

![rl1](https://user-images.githubusercontent.com/25974226/91436990-1acacb80-e8a4-11ea-9f0b-89874406c723.PNG)

@excel sheet

![rltkdcjd](https://user-images.githubusercontent.com/25974226/91434845-5cf20e00-e8a0-11ea-9ef7-27b55dc51401.PNG)
 
 
 
<Example 2>

@web page

![rl2](https://user-images.githubusercontent.com/25974226/91436997-1c948f00-e8a4-11ea-8461-1f1eaab75dc1.PNG)

@excel sheet

![dltkdgksxpdlqmf](https://user-images.githubusercontent.com/25974226/91434909-772bec00-e8a0-11ea-81d5-c6347a9cb743.PNG)



- Separation Excel View

![rltkdcjd1](https://user-images.githubusercontent.com/25974226/91434958-8b6fe900-e8a0-11ea-9bb7-20bce39cfcc0.PNG)


----------------------------------------------------------------------------------------------------------------------------

## 예시 Example 2 (Original, Version 1.1.3)

- 가로로 정렬된 개별 테이블을 하나의 테이블처럼 표시할 수 있습니다.

@web page

![2-1](https://user-images.githubusercontent.com/25974226/92365679-1ff30a80-f12f-11ea-9502-1a75d490c282.PNG)

@excel sheet

![2-2](https://user-images.githubusercontent.com/25974226/92365684-22556480-f12f-11ea-837b-3f814c56b4d6.PNG)



- 테이블 안에 중첩된 테이블을 별도로 표시할 수 있습니다.

@web page

![7-1](https://user-images.githubusercontent.com/25974226/92365689-24b7be80-f12f-11ea-82c7-6833d08ce0eb.PNG)

@excel sheet

![7-2](https://user-images.githubusercontent.com/25974226/92365692-25e8eb80-f12f-11ea-8ae5-2fa477379e9a.PNG)
