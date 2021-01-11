from docx.api import Document
#from xlsxwriter import Workbook
import pandas as pd 

path= "C:/Users/ygkwo/Desktop/test1/1.docx" #input_path
document=Document(path)
output_path="C:/Users/ygkwo/Desktop/test1" #output_path


#writer=pd.ExcelWriter('{}/docx_talbes.xlsx'.format(output_path), engine='xlsxwriter')



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

    df=pd.DataFrame(data)
    print(df)
    df.to_excel(writer,sheet_name='table_{}'.format(i))


#writer.save()