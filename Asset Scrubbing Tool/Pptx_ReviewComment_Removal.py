#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
from pptx import Presentation
import pptx
import re
import win32com.client


# In[ ]:


def reviewcommentsRemovalFromPPT(onlyfiles):
    
    global Corrupted_file
    Corrupted_file=[]
    
    for i in onlyfiles:
        try:
            ppt_app = win32com.client.GetObject(i)

            for ppt_slide in ppt_app.Slides:
                for comment in ppt_slide.Comments:
                    comment.delete()
                    ppt_app.Save()
                    ppt_app.Close()
        except:
            Corrupted_file.append(i)
            
    with open('Corrupted_Data_pptx.txt', 'w') as filehandle:
        filehandle.write("Issue in removing comments from pptx files")
        for items in Corrupted_file:
            filehandle.write('%s\n' % items)
    print("Comment Removal for PPTX completed")
            


# In[ ]:


# Folder_to_Scrub=r'C:\Users\vishal.a.kakkar\OneDrive - Accenture\Desktop\New Data\PPT'


# In[ ]:


# def getListOfFiles(Folder_to_Scrub):

#     listOfFile = os.listdir(Folder_to_Scrub)
#     allFiles = list()
#     # Iterate over all the entries
#     for entry in listOfFile:
#         # Create full path
#         fullPath = os.path.join(Folder_to_Scrub, entry)
#         # If entry is a directory then get the list of files in this directory 
#         if os.path.isdir(fullPath):
#             allFiles = allFiles + getListOfFiles(fullPath)
#         else:
#             allFiles.append(fullPath)

#     return allFiles

# global onlyfiles

# onlyfiles = getListOfFiles(Folder_to_Scrub)


# In[ ]:


# print(onlyfiles)


# In[ ]:


# reviewcommentsRemovalFromPPT(onlyfiles)


# In[ ]:


#def reviewcommentsRemovalFromPPT(Scrubbing_Location,Remove_Keywords,Scrubbing_pptxFileList):
#def reviewcommentsRemovalFromPPT(onlyfiles):
#     print("Removing reviews from ppt")
    
#     global Corrupted_Files
#     Corrupted_Files=[]
    
#     allSelectedFiles=[]
    
#     for File in onlyfiles:
#         allSelectedFiles.append(File)
        
#     for i in allSelectedFiles:
#         ppt_app = win32com.client.GetObject(i)
#         for ppt_slide in ppt_app.Slides:
#             for comment in ppt_slide.Comments:                        
#                 comment.delete()

