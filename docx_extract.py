from docx.api import Document
from xlsxwriter import Workbook
import pandas as pd 
from pandas import DataFrame
import io,sys

from folder_finding import find_docx_in_folder

# For test 
#path= "C:/Users/ygkwo/Desktop/test1/1.docx" #input_path
#output_path="C:/Users/ygkwo/Desktop/test1" #output_path


# Refrence (완성 시 삭제)
# one file
#docx_to_xlsx Function [Input, Output]
# def docx_to_xlsx(input,output):
#     document=Document(path)

#     writer=pd.ExcelWriter('{}/docx_talbes.xlsx'.format(output),engine='xlsxwriter')

#     for i in range(len(document.tables)):
#         table=document.tables[i]
#         data=[]
#         keys=None
#         row_data=None
        
#         for j,row in enumerate(table.rows):
#             text=(cell.text for cell in row.cells)
#             if j==0:
#                 keys=tuple(text)
#                 continue
#             row_data=dict(zip(keys,text))
#             data.append(row_data)
#             #
#             # df1=pd.DataFrame(data)
#             # print(df1)
#             # df1.to_excel(writer,sheet_name='table_{}'.format(i))

#         df=pd.DataFrame(data)
#         print(df)
#         df.to_excel(writer,sheet_name='table_{}'.format(i))


#     writer.save()


#multi
def docx_to_xlsx_multi(input):

    for k in input:
        print("file is "+ k)
        print("\n")

        document=Document(k)

        out=k[:-5]

        #checkpoint
        print("check point: slicing path is"+out)


        writer=pd.ExcelWriter('{}.xlsx'.format(out),engine='xlsxwriter')

        for i in range(len(document.tables)):
            table=document.tables[i]
            data=[]
            keys=None
            row_data=None
            
            for j,row in enumerate(table.rows):
                text=(cell.text for cell in row.cells)
                if j==0:
                    keys=tuple(text)
                    continue
                row_data=dict(zip(keys,text))
                data.append(row_data)
                #
                # df1=pd.DataFrame(data)
                # print(df1)
                # df1.to_excel(writer,sheet_name='table_{}'.format(i))

            df=pd.DataFrame(data)
            print(df)
            df.to_excel(writer,sheet_name='table_{}'.format(i))

        writer.save()


#excel 다른 칸으로 붙이는 법! 


##완성시 삭제!!! 

#one docx file
#docx_to_xlsx(path,output_path)

# print("Input Folder's PATH")
# input_path=input()

# print("Input_PATH is ",input_path)

# #If you want to reflect what you want to input_path, modify parameter

# #find_docx_in_folder(input_path)

# in_1=find_docx_in_folder(input_path)
# print("We'll extract this files(.docx)")
# print(in_1)


# docx_to_xlsx_multi(in_1)
