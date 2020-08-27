# HTML_Table_Excel
Scrapping HTML Table and Input a Table Data to Excel

- This Library apply HTML_Table_Extractor

- Library Name : Table_Excel
- Created Date : 27/Aug/2020
- Author : Minku Koo
- E-Mail : corleone@kakao.com
- Version : 1.0
- Keywords : 'Excel', 'Table', 'HTML', 'Crawling', 'Selenium', 'Extractor'



 * How to Use?
 ```
TableExcel = Table_Excel( URL_list <type=(String)list>, ChromeDriver Path <type=String>)
TableExcel.makeExel_abs( Excel File Path <type=String>, Table Header Color by Hex <type=String> (Default=F8E0EC) )
TableExcel.makeExel_sep( Excel File Path <type=String> )
```





 * Please, Import these Library : HTML_Table_Extractor, BeautifulSoup4, openpyxl, time, selenium
 * You should check your ChromeDriver exist
 * Also, You have to check, that your Chrome Version and your ChromeDriver version is same

----------------------------------------------------------------------------------------------------------------------------



- HTML의 table 태그의 데이터를 수집 및 변형하여 Excel 파일로 만들어주는 라이브러리 입니다.
- 엑셀 파일에는 링크, 페이지 제목이 포함되어 있습니다.
- 해당 웹 페이지의 모든 테이블을 수직으로 정렬시켜 표시합니다.
- 각 테이블의 헤더는 색을 달리하여 표시해줍니다.




** 사용법 **
 ```
TableExcel = Table_Excel( URL <리스트>, 크롬 드라이버 경로 <문자열>)
TableExcel.makeExel_abs( 엑셀 파일 경로 <문자열>, 테이블 헤더 색깔 - 16진수 <문자열> (Default=F8E0EC) )
TableExcel.makeExel_sep( 엑셀 파일 경로 <문자열> )
```




** 예시 Example **


- Absorption Excel View

<Example 1>
![rltkdcjd](https://user-images.githubusercontent.com/25974226/91434845-5cf20e00-e8a0-11ea-9ef7-27b55dc51401.PNG)
 
<Example 2>
![dltkdgksxpdlqmf](https://user-images.githubusercontent.com/25974226/91434909-772bec00-e8a0-11ea-81d5-c6347a9cb743.PNG)



- Separation Excel View
![rltkdcjd1](https://user-images.githubusercontent.com/25974226/91434958-8b6fe900-e8a0-11ea-9bb7-20bce39cfcc0.PNG)


