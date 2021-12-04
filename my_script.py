import streamlit as st
import pandas as pd
from marksheet_code import marksheet_rollwise,concise_marksheet,emails
import shutil
from zipfile import ZipFile
from PIL import Image
import os
import re

img=Image.open('iitplogo.png')
st.image(img)
st.title('Marksheet Generator')
st.header("Hii,Here you can get your latest quiz marksheets!!")

master_roll=pd.DataFrame()
responses=pd.DataFrame()

g = st.file_uploader("Upload the master_roll.csv file",type=['csv'])


if(g is not None):
    master_roll=pd.read_csv(g)
    st.write(master_roll.head())
    
else:
    st.warning("You need to upload a csv file")
    
    
n = st.file_uploader("Upload the response.csv file",type=['csv'])

if(n is not None):
    responses=pd.read_csv(n)
    st.write(responses.head())
    
else:
    st.warning("You need to upload a csv file")

correct = st.number_input("Enter the marks for each correct answer",0.000)
wrong = st.number_input("Enter the marks for each wrong answer",max_value=0.000)
if 'count' not in st.session_state:
	st.session_state.count = 0
if (g is not None and n is not None):        
    #st.write(result_1)
    st.write('')     
    if st.button("Generate Roll Number Wise Marksheet"):
      if os.path.isdir('marksheets'): 
        shutil.rmtree('marksheets')
      if not os.path.isdir("marksheets"):
        os.mkdir('marksheets')
      marksheet_rollwise(master_roll,responses,correct,wrong)
      st.session_state.count += 1
    
    #st.write(result_2)
    if st.button("Generate Concise Marksheet"):
      if not os.path.isdir("marksheets"):
        os.mkdir('marksheets')
      concise_marksheet(master_roll,responses,correct,wrong)
      path='marksheets/'
      df=pd.read_csv(path+'concise_marksheet.csv')
      st.write("Concise Marskheet:")
      st.write(df.head())
      st.download_button('Download Concise Marksheet', df.to_csv(), file_name='concise_marksheet.csv',mime="text/csv")
    if 'count2' not in st.session_state:
      st.session_state.count2=0
    result_3 = st.button("Send E-mail")
    if result_3:
      st.session_state.count2+=1
    if st.session_state.count2>0 and st.session_state.count>0:
      name=st.text_input("Enter your name:")
      user_name=st.text_input("Username:")
      pwd=st.text_input("Password :",type="password")
      if user_name is not None and pwd is not None:
        but=st.button("Submit")
        if but:
          emails(responses,name,user_name,pwd)
    else:
      if st.session_state.count==0:
        st.warning("Please generate marksheets first!!")
      else:
        st.write('')
    
else:
    result_1 = st.button("Generate Roll Number Wise Marksheet")
    #st.write(result_1)
    if result_1:
      st.warning("Please upload all CSV files")

    result_2 = st.button("Generate Concise Marksheet")
    #st.write(result_2)
    if result_2:
      st.warning("Please upload all CSV files")

    result_3 = st.button("Send Email")
    #st.write(result_3)
    if result_3:
      st.warning("Please upload all CSV files")