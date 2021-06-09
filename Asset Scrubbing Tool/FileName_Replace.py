#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import operator


# In[ ]:


def FileName_Replace(Scrubbing_Keywords,Scrubbing_FileList):
    
    for i in range(len(Scrubbing_FileList)):
        for a,b in Scrubbing_Keywords:
            if operator.contains(Scrubbing_FileList[i],a):
                txt=Scrubbing_FileList[i].replace(a, b)
                os.rename(Scrubbing_FileList[i],txt)
    print("File Name Replace Completed")


# In[ ]:


# Folder_to_Scrub=r'C:\Users\vishal.a.kakkar\OneDrive - Accenture\Desktop\BPD'

# Replace=[['VF', ''], ['VFC', ''], ['Kontoor', ''], ['Katalyst', ''], ['POLARIS', 'Client Vendor'],['TNF (A clothing store)','Client Vendor'],['Smart wool socks','Client Vendor'],['Omni','Client Vendor'],['Nora','NA'],['Vans','Client Vendor'],['RICEFW ID','####']]
# Replace=[['SWG','CLient'],['Southwest','Client']]


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

