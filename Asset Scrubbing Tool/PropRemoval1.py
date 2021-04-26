#!/usr/bin/env python
# coding: utf-8

# In[2]:



import os
import shutil
import zipfile

def prop_removal(folder_path):
    
#     folder_path = r"C:\Users\prasoon.raj\Desktop\Jupyter\Unscrubbed\Test_Property"
    for root, dirs, files in os.walk(folder_path):
            for file in files:
                try:
                    filename,fileext = os.path.splitext(file)
                    print("Property Removal : "+file)
                    directory_to_extract_to = root+"\\"+filename
                    file_path=os.path.join(root,file)

                    prop_path = directory_to_extract_to+"\docProps"

                    with zipfile.ZipFile(file_path, 'r') as zip_ref:
                        zip_ref.extractall(directory_to_extract_to)

                    os.remove(file_path)
                    
                    for file in os.listdir(prop_path):
                        os.remove(os.path.join(prop_path,file))

                    #Zipping and changing file back to ppt      
                    file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
                    fname,fext = os.path.splitext(file_path)
                    os.rename(file_path, fname + fileext)
                    shutil.rmtree(directory_to_extract_to)

                except Exception as e:
                    print(e)


# In[ ]:




