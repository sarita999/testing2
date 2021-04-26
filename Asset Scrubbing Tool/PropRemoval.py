#!/usr/bin/env python
# coding: utf-8

# In[1]:



import os
import shutil
import zipfile
from xml.dom import minidom

def prop_removal(folder_path):
    
#     folder_path = r"C:\Users\prasoon.raj\Desktop\Jupyter\Unscrubbed\Test_Property"
    for root, dirs, files in os.walk(folder_path):
            for file in files:
                try:
                    filename,fileext = os.path.splitext(file)
                    print("Property Removal : "+file)
                    directory_to_extract_to = root+"\\"+filename
                    file_path=os.path.join(root,file)

                    prop_path = directory_to_extract_to+"\docProps"

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
    #                                 print(tag.tagName)
                                    if fileext == ".xlsx":
                                        if (file == 'app.xml'):
                                            skiptag = ['Application','DocSecurity','ScaleCrop','LinksUpToDate','SharedDoc','HyperlinksChanged','AppVersion']
                                            if (xmltagName in skiptag):
        #                                         print("Skip" +tag.firstChild.nodeValue)
                                                continue
                                            else:
        #                                         print(tag.firstChild.nodeValue)
                                                tag.firstChild.nodeValue= " "
                                        elif (file == 'core.xml'):
                                            skiptag = ['cp:lastModifiedBy','dcterms:created','dcterms:modified','cp:lastPrinted']
                                            if (xmltagName in skiptag):
        #                                         print("Skip" +tag.firstChild.nodeValue)
                                                continue
                                            else:
        #                                         print(tag.firstChild.nodeValue)
                                                tag.firstChild.nodeValue= " "
                            
                                    else:
                                         tag.firstChild.nodeValue= " "
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


# In[3]:


# #!/usr/bin/env python
# # coding: utf-8

# # In[2]:



# import os
# import shutil
# import zipfile
# from xml.dom import minidom

# folder_path = r"C:\Users\prasoon.raj\Desktop\Jupyter\Unscrubbed\Test_Property"
# for root, dirs, files in os.walk(folder_path):
#             for file in files:
#                 try:
#                     filename,fileext = os.path.splitext(file)
#                     print("Property Removal : "+file)
#                     directory_to_extract_to = root+"\\"+filename
#                     file_path=os.path.join(root,file)

#                     prop_path = directory_to_extract_to+"\docProps"

#                     with zipfile.ZipFile(file_path, 'r') as zip_ref:
#                         zip_ref.extractall(directory_to_extract_to)

#                     os.remove(file_path)
#                     for file in os.listdir(prop_path):
#                         print(file)
#                         if file == "app.xml" or file == "core.xml" or file == "custom.xml":

#                             xmldoc = minidom.parse(os.path.join(prop_path,file))
#                             prop = xmldoc.firstChild

#                             for tag in xmldoc.firstChild.childNodes:
#                                 tagtype = type(tag.firstChild)
#     #                             print(tagtype)
#                                 xmltagName =tag.tagName
#                                 if tag.firstChild!=None:
#     #                                 print(tag.tagName)
#                                     if fileext == ".xlsx":
#                                         if (file == 'app.xml'):
#                                             skiptag = ['Application','DocSecurity','ScaleCrop','LinksUpToDate','SharedDoc','HyperlinksChanged','AppVersion']
#                                             if (xmltagName in skiptag):
#         #                                         print("Skip" +tag.firstChild.nodeValue)
#                                                 continue
#                                             else:
#         #                                         print(tag.firstChild.nodeValue)
#                                                 tag.firstChild.nodeValue= " "
#                                         elif (file == 'core.xml'):
#                                             skiptag = ['cp:lastModifiedBy','dcterms:created','dcterms:modified']
#                                             if (xmltagName in skiptag):
#         #                                         print("Skip" +tag.firstChild.nodeValue)
#                                                 continue
#                                             else:
#         #                                         print(tag.firstChild.nodeValue)
#                                                 tag.firstChild.nodeValue= " "
                            
#                                     else:
#                                          tag.firstChild.nodeValue= " "
#                             with open( os.path.join(prop_path,file), "w" ) as fs: 
#                                 fs.write( xmldoc.toxml() )
#                                 fs.close() 

#                     #Zipping and changing file back to ppt      
#                     file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
#                     fname,fext = os.path.splitext(file_path)
#                     os.rename(file_path, fname + fileext)
#                     shutil.rmtree(directory_to_extract_to)

#                 except Exception as e:
#                     print(e)


# In[ ]:


# import os
# import shutil

# directory_to_extract_to=r"C:\Users\prasoon.raj\Desktop\Jupyter\Unscrubbed\Test_Property\SRP_Wage Type Configuration Workbook ver 4_latest- 05th August - Copy"

# file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
# fname,fext = os.path.splitext(file_path)
# os.rename(file_path, fname + ".xlsx")
# shutil.rmtree(directory_to_extract_to)


# In[ ]:




