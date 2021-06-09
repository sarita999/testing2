#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import os
import docx
import re
import win32com.client as win32
from win32com.client import constants


# In[ ]:


def reviewcommentsRemovalFromDoc(onlyfiles):
    
    global Corrupted_Files
    Corrupted_Files=[]

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False 

    for i in onlyfiles:
        try:
            doc = word.Documents.Open(i) 
            doc.Activate()
            activeDoc = word.ActiveDocument
            if word.ActiveDocument.Comments.Count >= 1:
                word.ActiveDocument.DeleteAllComments()

            word.ActiveDocument.Save()
            doc.Close()
        except:
            Corrupted_Files.append(i)
            
    with open('Corrupted_Data_Docx.txt', 'w') as filehandle:
        filehandle.write("Issue in removing comments from docx files")
        for items in Corrupted_Files:
            filehandle.write('%s\n' % items)
    print("Comment Removal for Docx completed")


# In[ ]:


# Folder_to_Scrub=r'C:\Users\vishal.a.kakkar\OneDrive - Accenture\Desktop\BPD'


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


# reviewcommentsRemovalFromDoc(onlyfiles)


# In[ ]:


# print(Corrupted_Files)


# In[ ]:




