#!/usr/bin/env python
# coding: utf-8

# In[1]:


from docx import Document
import os
import docx
import re


# In[ ]:


def headerFooterRemoval(Scrubbing_Location,Scrubbing_docsFileList):
    print("Scrubbing Header and Footer Completely")

    
    def HeaderFooterRemovalFromDoc():
        for i in Scrubbing_docsFileList:
            if i.endswith(".docx"):
                document = Document(i)
                section = document.sections[0]

                header = section.header
                header.is_linked_to_previous = True

                footer = section.footer
                footer.is_linked_to_previous = True
                document.save(i)
        
    HeaderFooterRemovalFromDoc()
    
    

