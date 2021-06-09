#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tensorflow as tf
from tensorflow import keras
import matplotlib.pyplot as plt
import cv2
import os
from simple_image_download import simple_image_download as simp
import shutil
from PIL import Image, ImageGrab, ImageFilter
import io
import zipfile
import numpy as np
from tensorflow.keras.preprocessing.image import ImageDataGenerator
from tensorflow.keras.preprocessing import image
from tensorflow.keras.optimizers import RMSprop



Image.MAX_IMAGE_PIXELS = None
# warnings.simplefilter('ignore', Image.DecompressionBombWarning)


# In[ ]:





# In[ ]:





# In[ ]:




#--------Download from google images
def create_train_dataset():
    global source_dir
    global target_dir
    global file_names_src
    global file_names_des 
    global train_dataset
    global validation_dataset
    
    response = simp.simple_image_download
    response().download(txt+'-logo', 30)
    
    source_dir='./simple_images/'+txt+'-logo'
    target_dir ='./basedata/train/logo'
    
    file_names_src = os.listdir(source_dir)
    file_names_des = os.listdir(target_dir)
    
    for file_name in file_names_src:
        shutil.move(os.path.join(source_dir, file_name), target_dir)
    
    #----------Adding portion for using files inside Preferred Logo folder.
    
    PrefLogo_dir ='./Preferred_Logos'
    if os.path.exists(PrefLogo_dir):
        pass
    else:
        os.makedirs(PrefLogo_dir)

    files_preflogo = os.listdir(PrefLogo_dir)
    
    for file_name in files_preflogo:
        shutil.move(os.path.join(PrefLogo_dir, file_name), target_dir)
    
#     ----------------Creating rescaled pics------------------------------

#________________Test Code ____________Needs implementation in For loop____________

#     print(os.listdir(target_dir))
    for file_name in os.listdir(target_dir):
#         print(file_name)
        imgname,imgext = os.path.splitext(file_name)
        
        img = Image.open(os.path.join(target_dir, file_name))
        w,h = img.size

        w1 = int(w/2)
        h1= int(h/2)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_half"+imgext))
        
        w1 = int(w*2)
        h1= int(h*2)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_double"+imgext))
        w1 = int(w/3)
        h1= int(h/3)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_third"+imgext))
        w1 = int(w*3)
        h1= int(h*3)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_triple"+imgext))

        w1 = int(w/4)
        h1= int(h/4)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_fourth"+imgext))

        w1 = int(w*4)
        h1= int(h*4)

        # WIDTH and HEIGHT are integers
        resized_img = img.resize((w1, h1))
        resized_img.save(os.path.join(target_dir, imgname+"_quad"+imgext))


#     -----------------------------------------------------------------------------------------------------
    
    
    #-------------------------------------------------------    
    train = ImageDataGenerator(rescale= 1/255)
    validation = ImageDataGenerator(rescale= 1/255)

    train_dataset= train.flow_from_directory('./basedata/train/',
                                             target_size=(200,200),
                                             batch_size = 3,
                                            class_mode='binary')

    validation_dataset= validation.flow_from_directory('./basedata/validation/',
                                             target_size=(200,200),
                                             batch_size = 3,
                                            class_mode='binary')






# In[ ]:


#---Unzipping the document
def unzip_ppt(file_path):
    
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)

       
    count=0
    #Moving images to test folder
    for file in os.listdir(img_file_src):
        if(file.endswith(".jpg") or file.endswith(".png") or file.endswith(".jpeg")):
            shutil.copy(os.path.join(img_file_src,file),os.path.join(img_file_des,file))
        if(file.endswith(".emf")):
            count=count+1
    
    if(count!=0):
        file1 = open(r"./suspectEmf.txt","a")
        file1.write(os.path.basename(file_path)+"\n") 
        


# In[ ]:





# In[ ]:


def create_model():
    
    #model ---------------------------------------------------------------------------------
    
    model = tf.keras.models.Sequential([ tf.keras.layers.Conv2D(16,(3,3),activation = 'relu',input_shape=(200,200,3)),
                                    tf.keras.layers.MaxPool2D(2,2),
                                    #
                                    tf.keras.layers.Conv2D(32,(3,3),activation = 'relu'),
                                    tf.keras.layers.MaxPool2D(2,2),
                                    #
                                    tf.keras.layers.Conv2D(64,(3,3),activation = 'relu'),
                                    tf.keras.layers.MaxPool2D(2,2),
                                    ##
                                    tf.keras.layers.Flatten(),
                                    ##
                                    tf.keras.layers.Dense(512,activation= 'relu'),
                                    ##
                                    tf.keras.layers.Dense(1,activation='sigmoid')
                                    ])


    model.compile(loss='binary_crossentropy', 
                 optimizer = RMSprop(lr=0.001),
                 metrics =['accuracy'])
    
    
    return model
    
    
    


# In[ ]:


def train_model(model):
    checkpoint_dir = os.path.dirname(checkpoint_path+'/cp.ckpt')
    print(checkpoint_dir)
    cp_callback = tf.keras.callbacks.ModelCheckpoint(filepath=checkpoint_path+'/cp.ckpt',
                                                     save_weights_only=True,
                                                     verbose=1)

    model_fit = model.fit(train_dataset,
                         steps_per_epoch=21,
                         epochs= 60,
                         validation_data=validation_dataset,
                         callbacks=[cp_callback])
    return model


# In[ ]:


def load_model(checkpoint_path):
    latest = tf.train.latest_checkpoint(checkpoint_path)
    return latest


# In[ ]:


def start_img_scrub(f_path,client):
    
    global directory_to_extract_to
    global img_file_src
    global img_file_des
    global file_path
    global txt
    global checkpoint_path
    
# -----------Creating the basic folder structure needed for working with image scrubbing
    if os.path.exists('./basedata'):
        pass
    else:
        os.makedirs('./basedata/test')
        os.makedirs('./basedata/train/logo')
        os.makedirs('./basedata/train/not_logo')
        os.makedirs('./basedata/validation/logo')
        os.makedirs('./basedata/validation/not_logo')
        response = simp.simple_image_download
        response().download('blank-ppt-designs', 65)

        source_dir='./simple_images/blank-ppt-designs'

        file_names_src = os.listdir(source_dir)

        for file_name in file_names_src:
            shutil.move(os.path.join(source_dir, file_name), './basedata/train/not_logo')

        shutil.rmtree('./simple_images')

# ---------------------------------------------------------------------------------------------------------------
    folder_path =f_path

    txt=client
    txt=txt.lower()

    test_dir_path= 'basedata/test'

    checkpoint_path = 'Checkpoints/Checkpoint_'+txt
    print(checkpoint_path)
#---------------------------------------------------------------------------------
    # -----------------------------CLearing SuspectEmf file contents--------------
    if(os.path.exists(r"./suspectEmf.txt")):

        file = open(r"./suspectEmf.txt","r+")
        file.truncate(0)
        file.close()
#-------------------------------------------------------------
    #---------Creating a model---------------    
#-------------------------------------------------------------
    model = create_model()

    if(os.path.isdir(checkpoint_path)):

        latest = load_model(checkpoint_path)
        model.load_weights(latest).expect_partial()
    else:
        create_train_dataset()
        model = train_model(model)
#-------------------------------------------------------------
#-------------------------------------------------------------

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            try:
                filename,fileext = os.path.splitext(file)
                print(file)
                while filename[-1]==" ":
                    filename=filename[:-1]
                directory_to_extract_to = root+"\\"+filename
                file_path=os.path.join(root,file)

                if fileext == ".pptx":
                    img_file_src = directory_to_extract_to+"\ppt\media"
                elif fileext == ".docx":
                    img_file_src = directory_to_extract_to+"\word\media"
                elif fileext == ".xlsx":
                    img_file_src = directory_to_extract_to+"\\xl\media"    
                elif fileext == ".vsdx":
                    img_file_src = directory_to_extract_to+"\\visio\media"
                else :
                    continue

                img_file_des = r".\basedata\test"
                
                try:
                    unzip_ppt(file_path)
                except Exception:
                    errfile = open(r"./suspectEmf.txt","a")
                    errfile.write("Couldn't unzip file :"+file_path+"\n")
                    print("Couldn't unzip file")
                try:
                    os.remove(file_path)
                except Exception:
                    errfile = open(r"./suspectEmf.txt","a")
                    errfile.write("Couldn't delete file :"+file_path+"\n")
                    print("Couldn't delete file")
                
                #---------------------------------------------------------------------------------------------------------------
                for file in os.listdir(test_dir_path):
                    imgname,imgext = os.path.splitext(file)
                    img = image.load_img(test_dir_path+'//'+file,target_size=(200,200))
            #         print(file)
                    X= image.img_to_array(img)
                    X = np.expand_dims(X,axis=0)
                    images = np.vstack([X])

                    val = model.predict(images)

                    if val==0:
                        px_img = Image.open(os.path.join(test_dir_path,file))
            #  -------------------------------------------------------------------------------

                        img_width,img_height = px_img.size
                        result = Image.new('RGBA', (img_width, img_height), (255,255,255,0))
                        if imgext=='.jpeg' or imgext=='.jpg':
                            result = result.convert('RGB')
                        result.save(os.path.join(img_file_src,file))
                        px_img.close()


                #Zipping and changing file back to ppt
#                 output_path = r".\Scrub\\"+os.path.basename(root)+"\\"+filename       
                file_path=shutil.make_archive(directory_to_extract_to, 'zip', directory_to_extract_to)
                fname,fext = os.path.splitext(file_path)
                os.rename(file_path, fname + fileext)
                shutil.rmtree(directory_to_extract_to)


                #Removing 
                ###----Logo folder content
                logo_dir = './basedata/train/logo/'

                for file_name in os.listdir(logo_dir):
                    os.remove('./basedata/train/logo/'+file_name)
                ###----Test Folder content
                for file_name in os.listdir(test_dir_path):
                    os.remove('./basedata/test/'+file_name)

            except Exception as e:
                errfile = open(r"./suspectEmf.txt","a")
                errfile.write("Exception Occured in file :"+file_path+"\n")
                try:
                    shutil.rmtree(directory_to_extract_to)
                except:
                    print("Kindly Remove : " + directory_to_extract_to)
# In[ ]:



    


# In[ ]:



# start_img_scrub()


# print("Done")


# In[ ]:





# In[ ]:





# In[ ]:




