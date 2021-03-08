import io,sys
import os


#Get your PC's PATH
repositorypath=os.getcwd()
print(repositorypath)


sys.path.append(repositorypath)

from folder_finding import find_docx_in_folder
from docx_extract import docx_to_xlsx_multi
from docx_to_exist import docx_to_xlsx_exist

def main():

    #install essential package
    # os.system('pip install python-docx')
    # os.system('pip install xlsxwriter')
    # os.system('pip install pandas')
    # os.system('pip install workbook')
    # os.system('pip install openpyxl')
    
    #clear console
    os.system('cls')

    print("Document Automation Program")
    
    print("Before launching... ")
    print("If you open docx file, this program does not work.")
    #Input your folder PATH, then you can extract docx tables in your folder.
    print("Input Folder's PATH")
    input_path=input()
    
    #For test
    #input_path="C:\\Users\\ygkwo\\Desktop\\test1"
    

    print("Input_PATH is ",input_path)

    in_1=find_docx_in_folder(input_path)
    print("We'll extract this files(.docx)")
    print(in_1)

    docx_to_xlsx_multi(in_1)
    #docx_to_xlsx_exist(in_1)

if __name__=="__main__":
    main()

#Update 할것!! 

# 원하는 챕터에서의 테이블만 추출 
# #excel 다른 칸으로 붙이는 법!  