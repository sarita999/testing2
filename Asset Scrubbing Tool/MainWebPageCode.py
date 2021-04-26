#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, request, render_template
import os
import imgscrub
import PropRemoval
from imgscrub import *
import Main_Fn

app = Flask(__name__) 
#______________Default Page view--------------
@app.route('/')
def my_form():
    return render_template('scrub.html')

#________Page behaviour post scrub button is clicked.-----------
@app.route('/', methods=['POST'])
def my_form_post():
    #__Getting the folder path for the folder to scrub
    folder_path = request.form['fpath']
    scrub_path = './Scrub'
    
    print(folder_path)
    #___Getting the option for the type of scrubbing to be done_________
    scrubbing_selected = request.form['scrub']
    print(scrubbing_selected)
    if scrubbing_selected =='img':
        #___Getting Client Name_____
        clientName= request.form['clientName']
        
        imgscrub.start_img_scrub(folder_path,clientName)
        print("Image Scrubbing done with client name : "+clientName)
        
        try:
            PropRemoval.prop_removal(scrub_path)
        except Exception:
            print("Exception")
        
        
        
    elif scrubbing_selected =='txt':
        selected_txt_option= request.form['textscrub']
        print(selected_txt_option)
        try:
            PropRemoval.prop_removal(folder_path)
        except Exception:
            print("Exception")
            
        if selected_txt_option=='replacetext':
            keyword = request.form['replacekeyword']
            Flag_Status='replace'            
            Main_Fn.Main_function(folder_path,Flag_Status,keyword)
        else:
            keyword = request.form['removekeyword']
            Flag_Status='remove'            
            Main_Fn.Main_function(folder_path,Flag_Status,keyword)
            
        print(keyword)

    else:
        clientName= request.form['totalclientName']
        imgscrub.start_img_scrub(folder_path,clientName)
        
        print("Total Scrubbing selected with client name : "+clientName)
        selected_txt_option= request.form['totalscrub']
        print(selected_txt_option)
        if selected_txt_option=='replace':
            keyword = request.form['totalreplacekeyword']
            Flag_Status='replace'            
            Main_Fn.Main_function(scrub_path,Flag_Status,keyword)
        else:
            keyword = request.form['totalremovekeyword']
            Flag_Status='remove'            
            Main_Fn.Main_function(scrub_path,Flag_Status,keyword)
            
        print(keyword)
        try:
            PropRemoval.prop_removal(scrub_path)
        except Exception:
            print("Exception")
        
    return render_template('scrub.html')


app.run(debug=True,use_reloader=False)


# In[ ]:




