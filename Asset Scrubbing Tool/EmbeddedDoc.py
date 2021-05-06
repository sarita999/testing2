#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[ ]:


import os
import shutil
import zipfile
import imgscrub
import Main_Fn
import PropRemoval


def embedded_file_scrubbing(folder_path,scrub_type,keys):

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            print("Check for Embedded file on :"+ file)
            filename,fileext = os.path.splitext(file)
            directory_to_extract_to = root+"\\"+filename
            file_path=os.path.join(root,file)
    #         print(file_path)
            if fileext == ".pptx":
                embeded_file_path = directory_to_extract_to+"\ppt\embeddings"
            elif fileext == ".docx":
                embeded_file_path = directory_to_extract_to+"\word\embeddings"
            elif fileext == ".xlsx":
                embeded_file_path = directory_to_extract_to+"\\xl\embeddings"
            else :
                continue

            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(directory_to_extract_to)
            
            if os.path.isdir(embeded_file_path):
                print("Embedded files exist")
            
                try:
                    if scrub_type == 'img':
#                         print("Image Scrubbing")
                        imgscrub.start_img_scrub(embeded_file_path,keys)
                        print(keys)
                    elif scrub_type == 'txt_remove':
#                         print("Text Remove Scrubbing")
                        flag='remove'
                        Main_Fn.Main_function(embeded_file_path,flag,keys)
                    elif scrub_type == 'txt_replace':
#                         print("Text Replace Scrubbing")
                        flag='replace'
                        print(keys)
                        Main_Fn.Main_function(embeded_file_path,flag,keys)
                    elif scrub_type == 'both_remove':
#                         print("Image Scrubbing")
                        print(keys[0])
                        imgscrub.start_img_scrub(embeded_file_path,keys[0])
#                         print ("Text Remove scrubbing")
                        flag='remove'
                        print(keys[1])
                        Main_Fn.Main_function(embeded_file_path,flag,keys[1])
                    elif scrub_type == 'both_replace':
#                         print("Image Scrubbing")
                        print(keys[0])
#                         print ("Text Replace scrubbing")
                        imgscrub.start_img_scrub(embeded_file_path,keys[0])
                        flag='replace'
                        print(keys[1])
                        Main_Fn.Main_function(embeded_file_path,flag,keys[1])
            
                    PropRemoval.prop_removal(embeded_file_path)

                except Exception as e:
                    print("Exception occured while scrubbing embedded files")
                    print(e)
                   
                scrub_file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
                fname,fext = os.path.splitext(scrub_file_path)
            
                print(file_path)
                os.remove(file_path)
                os.rename(scrub_file_path, fname + fileext)
                print("Done with Embeded files")
                shutil.rmtree(directory_to_extract_to)
            else:
                print("Couldn't find embedded files")
                shutil.rmtree(directory_to_extract_to)
                


# In[ ]:




