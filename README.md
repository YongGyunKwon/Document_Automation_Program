# 문서 자동화 프로그램 (Document Automation Pragram)

### 사용목적(Purpose)

`사양서를 수동으로 작성하는 도중에 발생할 수 있는 "인간의 실수"를 방지하기 위해 본 프로그램을 만든다!`

We create this program to preven "Human Error" that can occur while manually writing specifications!

### 사용언어(language)

- Python

### 사용 라이브러리(Using library)

- io
- sys
- os
- python-docx

    ```tsx
    pip install python-docx
    ```

- xlsxwriter

    ```tsx
    pip install xlsxwriter
    ```

- pandas

    ```tsx
    pip install pandas
    ```

- workbook

    ```tsx
    pip install workbook
    ```

### 작동과정

1. 워드 파일을 선택한다. 

    (Select word file)

2. 워드 문서 내의 테이블들을 찾아서 추출한다. 

    (Extract table in selected word file)

3. 추출한 테이블을 이용하여 새로운 엑셀 파일을 만들어서 붙여넣는다. 

    (Using extracted table, make new Excel file.)

### How to use

     

*[사용전에!] 추출하려는 docx파일이 열려있으면, 프로그램이 제대로 작동하지 않는다.

1. `main.py`를 실행한다.
2. `Input Folder's PATH` 가 나온다면 docx파일이 있는 경로를 복사한다.
3. 폴더안의 .docx 파일들 안의 테이블(표)을 인식하여 같은 이름의 .xlsx 파일을 생성된다.

    ```tsx
    ex)
    1.docx 안에 테이블(표)이 1.xlsx라는 파일로 만들어 진다.
    (.docx의 테이블들은 각각의 sheet으로 분리되며  만들어지며 table_0 , table_1 순서대로 만들어진다.)
    ```

### Update

- Repository 경로 자동 지정
- 필수 라이브러리 자동설치