#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import shutil
import zipfile
from xml.dom import minidom


def prop_removal(folder_path):
    
    tags_to_scrub = ['Template','dc:creator','Company','Manager','dc:subject','dc:title','cp:keywords','dc:description','cp:category']
    formats_sup = ['.pptx','.xlsx','.docx']
    for root, dirs, files in os.walk(folder_path):
                for file in files:
                    try:
                        filename,fileext = os.path.splitext(file)
                        print("Property Removal : "+file)
                        directory_to_extract_to = root+"\\"+filename
                        file_path=os.path.join(root,file)

                        prop_path = directory_to_extract_to+"\docProps"
                        
                        if fileext not in formats_sup:
                            continue
                        
                        with zipfile.ZipFile(file_path, 'r') as zip_ref:
                            zip_ref.extractall(directory_to_extract_to)

                        os.remove(file_path)
                        for file in os.listdir(prop_path):
                            print(file)
                            if file == "app.xml" or file == "core.xml" or file == "custom.xml":

                                xmldoc = minidom.parse(os.path.join(prop_path,file))
                                prop = xmldoc.firstChild

                                for tag in xmldoc.firstChild.childNodes:
                                    tagtype = type(tag.firstChild)
        #                             print(tagtype)
                                    xmltagName =tag.tagName
                                    if tag.firstChild!=None:
                                        if (xmltagName in tags_to_scrub):
#                                             print(tag.tagName)
#                                             print(tag.firstChild.nodeValue)
                                            tag.firstChild.nodeValue= " "
                                        else:
                                            continue
                                with open( os.path.join(prop_path,file), "w" ) as fs: 
                                    fs.write( xmldoc.toxml() )
                                    fs.close() 

                        #Zipping and changing file back to ppt      
                        file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
                        fname,fext = os.path.splitext(file_path)
                        os.rename(file_path, fname + fileext)
                        shutil.rmtree(directory_to_extract_to)

                    except Exception as e:
                        print(e)


# In[ ]:





# In[ ]:




