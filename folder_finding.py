import os
 

#path: folder's path
def find_docx_in_folder(path):
    os.chdir(path)

    docx_files=[]

    for files in os.listdir("."):
        #확장자 더 찾아보기 (docx 외에)
        if files.endswith(".docx") or files.endswith(".doc") or files.endswith(".docm") or files.endswith(".dot") or files.endswith(".dotm") or files.endswith(".dotx"):
            print(files)
            docx_file_path=path+'\\'+files
            print(docx_file_path)
            docx_files.append(docx_file_path)    
        


    return docx_files


# For test
#print("Input path(folder)")


##완성하면 삭제하기!!! 

#input
#input_path=input()

#input_path.replace('\'','/')

#print("Input_PATH is ",input_path)


#print(find_docx_in_folder(input_path))


#test1=find_docx_in_folder(input_path)

##