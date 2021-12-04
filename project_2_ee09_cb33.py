import streamlit as st
import os
import csv
import pandas as pd
#from tqdm import tqdm
from fpdf import FPDF
from PIL import Image
from pdf_code import transcript_generator
import numpy as np
#import matplotlib.pyplot as plt
import shutil
from zipfile import ZipFile

img=Image.open('iitplogo.png')
st.image(img)
st.title('Transcript Generator')
st.header("Hey mate!! You can now get your transcript easily by uploading the required files!!")
st.write('')
st.write('')
st.header("Mandatory Requirements")
grad=pd.DataFrame()
sub_master=pd.DataFrame()
names_with_roll=pd.DataFrame()
csv1 = st.file_uploader("Upload the grades.csv file",type=['csv'])
if(csv1 is not None):
    grad=pd.read_csv(csv1)
    st.write("This is what it looks like : ")
    st.dataframe(grad)
    
else:
    st.warning("You need to upload a csv file!!")
    
    
csv2 = st.file_uploader("Upload the subjects_master.csv file",type=['csv'])
if(csv2 is not None):
    sub_master=pd.read_csv(csv2)
    st.write("This is what it looks like : ")
    st.dataframe(sub_master)
    
else:
    st.warning("You need to upload a csv file!!")
    
       
csv3 = st.file_uploader("Upload the names_roll.csv file",type=['csv'])

if(csv3 is not None):
    names_with_roll=pd.read_csv(csv3)
    st.write("This is what it looks like : ")
    st.dataframe(names_with_roll)
    
else:
    st.warning("You need to upload a csv file!!")
    
st.write('')   
st.header("Optional requirements:")
stamp=st.file_uploader("Upload the image of seal",type=["png","jpeg","jpg"])
if(stamp is not None):
    img = Image.open(stamp)
    img1=img.save('stamp_iitp.png')
    #x=cv2.imwrite(r'stamp_iitp.png',stamp)
    st.write("Image uploaded successfully!!")
    st.image(img)
sign=st.file_uploader("Upload the image of signature",type=["png","jpeg","jpg"])
if(sign is not None):
    img2 = Image.open(sign)
    img22=img2.save('assistant_reg.png')
    st.write("Image uploaded successfully!!") 
    st.image(img2)
#user_input = st.text_input("Enter the range of roll numbers for which you want to generate transcript")
start_roll=st.text_input("Enter the starting roll number of the range")
end_roll=st.text_input("Enter the ending roll number of the range")

if(csv1 is not None and csv2 is not None and csv3 is not None):
    if 'count1' not in st.session_state:
            st.session_state.count1=0
    if 'invalid_rolls' not in st.session_state:
        st.session_state.invalid_rolls=[]
    if(len(start_roll)>1 and len(end_roll)>1 and start_roll[0:6]==end_roll[0:6]):
        start_roll=start_roll.upper()
        end_roll=end_roll.upper()
        rolls=[]
        r=int(end_roll[6:])-int(start_roll[6:])+1
        for i in range(r):
            if(int(start_roll[6:])+i<10):
                rolls.append(start_roll[0:6]+"0"+str(int(start_roll[6:])+i))
            else:
                rolls.append(start_roll[0:6]+str(int(start_roll[6:])+i))
        st.write("Range entered succesfully")
        if st.button('Generate in-range Transcripts'):
            if os.path.isdir('transcriptsIITP'):
                shutil.rmtree('transcriptsIITP')
            os.mkdir('transcriptsIITP')
            st.write("Generating transcripts...")
            invalid_rolls=transcript_generator(grad,sub_master,names_with_roll,rolls,stamp,sign)
            st.session_state.count1+=1
            st.write("All transcripts successfully generated!!")
            if not invalid_rolls:
                st.session_state.invalid_rolls=["All roll numbers exist.Nothing to show!!"]
            else:
                st.session_state.invalid_rolls=invalid_rolls           
            
            zipObj = ZipFile('transcripts.zip', 'w')
            path='transcriptsIITP/'
            
            for k in os.listdir(path):
                zipObj.write(path+k)
                
            zipObj.close()
            
    else:
        if st.button("Generate in-range Transcripts"):
            st.warning("Please enter correct range of roll numbers.Range entered should be of same year,same programme and same branch!!")
    if st.session_state.invalid_rolls:
        if(st.button("Non-existent Roll Numbers")):
            st.write(st.session_state.invalid_rolls)
    if 'count2' not in st.session_state:
            st.session_state.count2=0
    if st.button("Generate all Transcripts"):
        st.session_state.invalid_rolls=[]
        if os.path.isdir('transcriptsIITP'):
            shutil.rmtree('transcriptsIITP')
        os.mkdir('transcriptsIITP')
        rolls=list(names_with_roll.iloc[:,0])
        #print(rolls)
        st.write("Generating transcripts...")
        invalid_rolls=transcript_generator(grad,sub_master,names_with_roll,rolls,stamp,sign)
        st.write("All transcripts successfully generated!!")
        st.session_state.count2+=1
        
        zipObj = ZipFile('transcripts.zip', 'w')
        path='transcriptsIITP/'
        
        for k in os.listdir(path):
            zipObj.write(path+k)
            
        zipObj.close()
    if st.session_state.count1>0 or st.session_state.count2>0:
        with open('transcripts.zip', 'rb') as f:
            st.download_button('Download Zip', f, file_name='transcripts.zip')
    if os.path.isfile('stamp_iitp.png'):
        os.remove('stamp_iitp.png')
    if os.path.isfile('assistant_reg.png'):
        os.remove('assistant_reg.png')
else:
    if st.button("Generate in-range Transcripts"):
        st.warning("Please upload all the mandatory files.")
    if st.button("Generate all Transcripts"):
        st.warning("Please upload all the mandatory files.")