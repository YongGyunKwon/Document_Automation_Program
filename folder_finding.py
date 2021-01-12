import os 

#for test
path1="C:/Users/ygkwo/Desktop/test2"


#original
# os.chdir(path)

# for files in os.listdir("."):
#     if files.endswith(".docx"):
#         print(files)

#file list in multi folder


def find_docx_in_folder(path):
    os.chdir(path)

    docx_files=[]

    for files in os.listdir("."):
        if files.endswith(".docx"):
            print(files)
            docx_file_path=path+'\\'+files
            print(docx_file_path)
            docx_files.append(docx_file_path)    
            
    return docx_files

print("Input path(folder)")

#input
input_path=input()

#input_path.replace('\'','/')

print("Input_PATH is ",input_path)


#If you want to reflect what you want to input_path, modify parameter

#find_docx_in_folder(input_path)



print(find_docx_in_folder(input_path))