import io,sys

sys.path.append("C:/Users/ygkwo/Desktop/docxtool/Document_Automation_Program")

from folder_finding import find_docx_in_folder
from docx_extract import docx_to_xlsx_multi

def main():
    print("Main")
    
    print("Input Folder's PATH")
    input_path=input()

    print("Input_PATH is ",input_path)

    in_1=find_docx_in_folder(input_path)
    print("We'll extract this files(.docx)")
    print(in_1)

    docx_to_xlsx_multi(in_1)


if __name__=="__main__":
    main()