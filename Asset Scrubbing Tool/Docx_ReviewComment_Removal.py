#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from docx import Document
import os
import docx
import re
import win32com.client as win32


# In[ ]:


def reviewcommentsRemovalFromDoc(Scrubbing_Location,Remove_Keywords,Scrubbing_docsFileList):
   allSelectedFiles=[]

   for i in Scrubbing_docsFileList:
        if i.endswith(".docx"):
           path_file_name = i
           win32.pythoncom.CoInitialize ()
           word = win32.gencache.EnsureDispatch("Word.Application")
           word.Visible = False
           doc = word.Documents.Open(path_file_name)
           doc.Activate()
           word.ActiveDocument.TrackRevisions = False  # Maybe not need this (not really but why not)

           # Accept all revisions
           word.ActiveDocument.Revisions.AcceptAll()
           # Delete all comments
           if word.ActiveDocument.Comments.Count >= 1:
               word.ActiveDocument.DeleteAllComments()

           word.ActiveDocument.Save()
           doc.Close(False)
           word.Application.Quit()
   print("Reviews scrubbed from the doc files")

