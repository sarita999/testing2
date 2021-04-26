#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import os
import docx
import re


# In[ ]:


def FileName_Replace(Scrubbing_Location,Scrubbing_Keywords,Scrubbing_FileList):
    print("FileName_Replace Called")
    
    global dictOfReplacedInput
    dictOfReplacedInput= dict(x.split(':') for x in Scrubbing_Keywords.split(','))
    print(dictOfReplacedInput)
    
    allSelectedFiles=[]
    
    for File in Scrubbing_FileList:
        allSelectedFiles.append(File)

    
    filenameContainList = []
    updatedFileNameList = []
    removeList1 = []
    
    
    #creating path in correct format
    for i in allSelectedFiles:
        x = i.replace('/', '\\')
        removeList1.append(x)
        

        
    #making replacing path in correct format
    FileLocation1 = Scrubbing_Location.replace('/','\\\\')   
    
    #seperating key value elements
    key, value = [], []
    for k, v in dictOfReplacedInput.items():
        key.append(k)
        value.append(v)

    
    
    #getting all filename from path from last
    lastFileName = [x.split('\\')[-1] for x in removeList1]
    
    fileNames, fileType =[], []
    for z in lastFileName:
        fname,fext = os.path.splitext(z)
        fileNames.append(fname)
        fileType.append(fext)
    
    
    #getting all file names with userInput after matching with user input
    for i in key:
        matching = list(s for s in fileNames if i.lower() in s.lower())
        filenameContainList.extend(matching)


    #removing the user input from file name using replacing logic
    for k,v in dictOfReplacedInput.items():
        repat = "(.*){}(.*)".format(k)
        for j in filenameContainList:
            tmp = re.search(repat, j, re.IGNORECASE)
            if tmp:
                token = v.join(tmp.groups())
                updatedFileNameList.append(token)
    
    
    #after getting source and destination location in correct format,renaming the file
    count = 0
    
    for i in filenameContainList:
        x = Scrubbing_Location + '\\' + lastFileName[count]
        x1 = x.replace('\\', '\\\\')
        y = Scrubbing_Location + '\\' + updatedFileNameList[count] +fileType[count]
        y1 = y.replace('\\', '\\\\')
        try:
            os.rename(x1,y1)
            count = count + 1
        except:
            print("Couldn't rename the file")
            


# In[ ]:




