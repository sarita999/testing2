#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import os
import docx
import re


# In[ ]:


def FileName_Removal(Scrubbing_Location,Scrubbing_Keywords,Scrubbing_FileList):
    print("FileName_Removal Called")
    
    
    def funcCreateListRemoveChar():
        global userInputForRemove
        userInputForRemove = []
        var2 = ''.join(Scrubbing_Keywords.split(" "))
        list1 = var2.split(',')
        for elements in list1:
            userInputForRemove.append(elements.lower())
    funcCreateListRemoveChar()
    

    
    allSelectedFiles=[]
    
    for File in Scrubbing_FileList:
        allSelectedFiles.append(File)
        

        
    
    filenameContainList = []
    updatedFileNameList = []
    removeList1 = []
    
    
    #using list comprehension to concat the selected items
    #allSelectedFiles = [y for x in [selectedFilesFN, selectAllFilesFN] for y in x]
    
    
    #creating path in correct format
    for i in allSelectedFiles:
        x = i.replace('/', '\\')
        removeList1.append(x)
        

              
    #replacing path in correct format
    FileLocation1 = Scrubbing_Location.replace('/','\\\\')
    

    
    #getting all filename from path from last
    lastFileName = [x.split('\\')[-1] for x in removeList1]
    
    
    fileNames, fileType =[], []
    for z in lastFileName:
        fname,fext = os.path.splitext(z)
        fileNames.append(fname)
        fileType.append(fext)
    
    
    #getting all file names with userInput after matching with user input
    for i in userInputForRemove:
        matching = list(s for s in fileNames if i.lower() in s.lower())
        filenameContainList.extend(matching)


    replace_ = re.compile("|".join(userInputForRemove), flags=re.IGNORECASE)
    updatedFileNameList = [replace_.sub("", x) for x in fileNames]
    
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
            print("Couldn't rename the file : ",i)
    


# In[ ]:




