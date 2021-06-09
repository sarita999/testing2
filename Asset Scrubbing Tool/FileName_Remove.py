#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import operator


# In[ ]:


def FileName_Removal(Scrubbing_Keywords,Scrubbing_FileList):
    
    for i in range(len(Scrubbing_FileList)):
        for a in Scrubbing_Keywords:
            if operator.contains(Scrubbing_FileList[i],a):
                txt=Scrubbing_FileList[i].replace(a,"")
                os.rename(Scrubbing_FileList[i],txt)
    print("File Name Remove Completed")

