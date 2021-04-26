#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
from os.path import isfile, join
import Docx_Fn
import Vsdx_Module
import Xlsx_Module
import Pptx_Fn


# In[ ]:


def Main_function(folder_path,Flag_Status,keyword):
    class scrubbing(object):
        def __init__(self,Location,Keywords,Flag):
            self.Location=Location
            self.Keywords=Keywords
            self.Flag=Flag

        def Files_in_Folder(self):

            self.File_Seperation()

        def File_Seperation(self):
            global docsFileList
            global xlsxFileList
            global pptxFileList
            global csvFileList
            global txtFileList
            global vsdxFileList

            docsFileList = []
            xlsxFileList = []
            pptxFileList = []
            csvFileList = []
            txtFileList = []
            vsdxFileList = []


            for i in onlyfiles:
                if i.endswith(".docx"):
                    docsFileList.append(i)
                elif i.endswith(".xlsx"):
                    xlsxFileList.append(i)
                elif i.endswith(".pptx"):
                    pptxFileList.append(i)
                elif i.endswith(".vsdx"):
                    vsdxFileList.append(i)
                elif i.endswith(".vsd"):
                    vsdxFileList.append(i)

            self.Calling_Fn()

        def Calling_Fn(self):
            if (len(docsFileList)) >0:
                Docx_Fn.Scrubbing_Process(self.Location,self.Keywords,self.Flag,docsFileList)
            if (len(xlsxFileList)) >0:
                Xlsx_Module.Scrubbing_Process_xlsx(self.Location,self.Keywords,self.Flag,xlsxFileList)
            if (len(pptxFileList)) >0:
                Pptx_Fn.Scrubbing_Process(self.Location,self.Keywords,self.Flag,pptxFileList)
            if (len(vsdxFileList)) >0:
                Vsdx_Module.Scrubbing_Process_vsdx(self.Location,self.Keywords,self.Flag,vsdxFileList)               


    #Folder_to_Scrub=input("Enter the location of the Folder to scrub--")
    Folder_to_Scrub=folder_path
    #Folder_to_Scrub=r"C:\Users\vishal.a.kakkar\OneDrive - Accenture\Desktop\MyConcerto\Scrubbing Tool\Anujs Work\Files\New folder (2)"
    #Flag=input("Enter the operation you want to perform 'Remove' or 'Replace' --")
    Flag=Flag_Status
    Flag=Flag.lower()
    
    def getListOfFiles(Folder_to_Scrub):

        listOfFile = os.listdir(Folder_to_Scrub)
        allFiles = list()
        # Iterate over all the entries
        for entry in listOfFile:
            # Create full path
            fullPath = os.path.join(Folder_to_Scrub, entry)
            # If entry is a directory then get the list of files in this directory 
            if os.path.isdir(fullPath):
                allFiles = allFiles + getListOfFiles(fullPath)
            else:
                allFiles.append(fullPath)

        return allFiles

    
    global onlyfiles
    onlyfiles = getListOfFiles(Folder_to_Scrub)


    if (Flag=="replace"):
        #Replace_Keywords=input("Enter the keywords and the text you want to replace it with--")
        Folder=scrubbing(Folder_to_Scrub,keyword,Flag)
    elif (Flag=="remove"):
        #Keywords_to_Scrub=input("Enter the Keywords that you want to scrub--")
        Folder=scrubbing(Folder_to_Scrub,keyword,Flag)
    else:
        print("INCORRECT INPUT")



    Folder.Files_in_Folder()


# In[ ]:




