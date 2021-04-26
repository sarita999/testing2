#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
from pptx import Presentation
import pptx
import re
import win32com.client as win32
import win32com.client


# In[ ]:


def reviewcommentsRemovalFromPPT(Scrubbing_Location,Remove_Keywords,Scrubbing_pptxFileList):
    print("Removing reviews from ppt")
    allSelectedFiles=[]
    
    for File in Scrubbing_pptxFileList:
        allSelectedFiles.append(File)
        
    for i in allSelectedFiles:
        if i.endswith(".pptx"):
            #ppt_dir = i
            ppt_app = win32com.client.GetObject(os.path.abspath(i))
            for ppt_slide in ppt_app.Slides:
                for comment in ppt_slide.Comments:    
                    print(comment.Text)
                    comment.delete()


# In[ ]:




