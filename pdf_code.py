# -*- coding: utf-8 -*-
"""Project 2 (EE09, CB33)

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/drive/1PU0kzkWdh3raeL_-AQsizvEDGRGFXzJr
"""

#from google.colab import drive
#drive.mount('/content/drive')

#cd /content/drive/MyDrive/tuts_2021/project 2

#!pip install fpdf

import os
import csv
import pandas as pd
from fpdf import FPDF
import streamlit as st
from datetime import datetime
from pytz import timezone

def cpi_calc(grades,subjects_master,names_roll):
  grades_to_number={'AA':10,'AB':9,'BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'F':0,'I':0}
  marksheet={}

  #with open('grades.csv', newline='') as grades:
  for index,line in grades.iterrows():
      #print(line)
      if line['Roll'] not in marksheet.keys(): 
          marksheet[line['Roll']]={}
          marksheet[line['Roll']][line['Sem']]=[]
          marksheet[line['Roll']][line['Sem']].append((line['Grade'],line['SubCode'],line['Sub_Type']))
      else:
          if(line['Sem'] not in marksheet[line['Roll']].keys()):
              marksheet[line['Roll']][line['Sem']]=[]
              marksheet[line['Roll']][line['Sem']].append((line['Grade'],line['SubCode'],line['Sub_Type']))
          else:
              marksheet[line['Roll']][line['Sem']].append((line['Grade'],line['SubCode'],line['Sub_Type']))

  subject={}
  ltp={}
  credits={}
  #with open('subjects_master.csv', newline='') as subjects_master:
  #sheet1 = csv.DictReader(subjects_master)
  for i,line in subjects_master.iterrows():
      subject[line['subno']]=line['subname']
      ltp[line['subno']]=line['ltp']
      credits[line['subno']]=line['crd']

  roll_names={}
  branch={}
  #with open('names-roll.csv', newline='') as names_roll:
  #sheet2 = csv.DictReader(names_roll)
  for i,line in names_roll.iterrows():
      roll_names[line['Roll']]=line['Name']
      branch[line['Roll']]=line['Roll'][4:6]

  overall={}
  sem_result={}
  i=0
  #credits_cleared={}
  for roll,sems in marksheet.items():
      for sem_no,lis in sems.items():
          grade = [i[0] for i in lis]
          sub_type = [i[2] for i in lis]
          sub_code = [i[1] for i in lis]          
          sub_credit=[credits[i] for i in sub_code]
          grade_total=0
          clr_crd=0
          for i in range(len(grade)):
              grade[i]=grade[i].strip()
              
              if(grade[i][-1]=='*'):
                  grade[i]=grade[i][:-1]
                  
              grade_total+=grades_to_number[grade[i]]*(int)(sub_credit[i])
          credit_total=0
          for i in sub_credit:
            credit_total+=(int)(i)   
          spi=round(grade_total/credit_total,2)
          if(roll not in overall.keys()):
              overall[roll]=[]
              sem_result[roll]=[]
              
          sem_result[roll].append((sub_code,sub_type,grade))  
          overall[roll].append((credit_total,spi))
          
  sem_credits={}
  spi={}
  CPI={}
  total_credits_taken={}
  big_data={}
  credits_cleared={}
  for roll,roll_item in overall.items():
      sem_credits[roll]=[]
      spi[roll]=[]
      for i in roll_item:
        sem_credits[roll].append((int)(i[0]))
        spi[roll].append(i[1])
      
      total_credits_taken[roll]=[]
      total=0
      c=0
      CPI[roll]=[]
      for i in range(len(sem_credits[roll])):
          total+=sem_credits[roll][i]
          c+=sem_credits[roll][i]*spi[roll][i]
          CPI[roll].append(round(c/total,2))
          total_credits_taken[roll].append(total)    
      big_data[roll]=[]
      credits_cleared[roll]=[]
      for i in range(len(sem_credits[roll])):
          idx=0
          sem_grades=[]
          sub_code=[]
          sub_name=[]
          l_t_p=[]
          cred=[]
          grad=[]
          sub_code.append("Sub. Code")
          sub_name.append("Subject Name")
          l_t_p.append("L-T-P")
          cred.append("CRD")
          grad.append("GRD")
          crd_grd=0
          for j in range(len(sem_result[roll][i][2])):
              v=sem_result[roll][i]
              sub_code.append(v[0][j])
              sub_name.append(subject[v[0][j]])
              l_t_p.append(ltp[v[0][j]])
              cred.append(credits[v[0][j]])
              grad.append(v[2][j])
              if(v[2][j]=='F' or v[2][j]=='I'):
                crd_grd+=0
              else:
                crd_grd+=credits[v[0][j]]
          #print(roll,i+1,crd_grd)
          credits_cleared[roll].append(crd_grd)
          sem_grades.append(sub_code)
          sem_grades.append(sub_name)
          sem_grades.append(l_t_p)
          sem_grades.append(cred)
          sem_grades.append(grad)
          big_data[roll].append(sem_grades)
  
  return big_data,CPI,spi,sem_credits,credits_cleared,roll_names
      #print(big_data[roll])

def transcript_generator(grad,sub_mast,names,rolls,stamp,sign):
  class PDF(FPDF):
    w=210
    h=297
    def lines(self):
      self.set_line_width(1.0)
      self.line(0, h/2, 210, h/2)
      self.rect(8.4, 12, 200.0,287.0)

  big_data,CPI,spi,sem_credits,credits_cleared,roll_names=cpi_calc(grad,sub_mast,names)
  branch_name={"CS":"Computer Science and Engineering","EE":"Electrical Engineering","ME":"Mechanical Engineering","CE":"Civil and Environmental Enngineering","CB":"Chemical and Biochemical Engineering"}
  invalid_rolls=[]
  path="transcriptsIITP"
  if not os.path.isdir(path):
    os.mkdir(path)
  for r in rolls:
    r=r.upper()
    if(r in big_data.keys()):
      w=h=0
      year="20"+str(r[0:2])
      if(r[2]=='0'and r[3]=='1'):
        #creating PDF file
        pdf_BTech = PDF(orientation='L', unit='mm', format='A3')
        w=420
        h=297
        pdf_BTech.add_page()
        #Images
        pdf_BTech.set_xy(0.12*w/6+w/50, w/50+(h-2*(w/50))*0.1*0.1)
        pdf_BTech.image('iitp-1_black.jpeg',  link='', type='', w=0.08*w, h=(h-2*(w/50))*0.1*0.7)
        pdf_BTech.set_xy(w-0.14*w+0.12*w/6, w/50+(h-2*(w/50))*0.1*0.1)
        pdf_BTech.image('iitp-1_black.jpeg',  link='', type='', w=0.08*w, h=(h-2*(w/50))*0.1*0.7)
        pdf_BTech.set_xy(0.14*w, w/50)
        pdf_BTech.image('iitp_heading.png',  link='', type='', w=282, h=24.3)
        if stamp is not None:
          pdf_BTech.set_xy(0.38*w, 0.7*h)
          pdf_BTech.image('stamp_iitp.png',  link='', type='', w=40.03, h=40.48)
        if sign is not None:
          pdf_BTech.set_xy(0.825*w, 0.725*h)
          pdf_BTech.image('assistant_reg.png',  link='', type='', w=40.03, h=21.08)

        #Text
        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.set_font('Arial', 'B', 10)
        pdf_BTech.set_text_color(0, 0, 0)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Roll No.:", border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+5, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=str(r), border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h/2)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Programme:", border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+5, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h/2)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='C', txt="Bachelor of Technology", border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+0.6*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Name:", border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+0.6*(w-2*w/50)/3+5, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=roll_names[r], border=0)

        #add name=roll_names[r]

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+0.6*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h/2)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Course:", border=0)

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+0.6*(w-2*w/50)/3+5, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h/2)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=branch_name[str(r[4:6])], border=0)
        #add branch name=branch_name[branch[r]]

        pdf_BTech.set_xy(0.2*(w-2*w/50)+w/50+5+2*0.6*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='C', txt="Year of Admission: "+year, border=0)

        pdf_BTech.set_font('Arial', 'B', 12)
        pdf_BTech.set_xy(0.82*w, 0.85*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Assistant Registrar (Academic)", border=0)

        #date
        now=datetime.now()
        ind_time = datetime.now(timezone("Asia/Kolkata")).strftime('%d %b %Y, %H:%M')
        pdf_BTech.set_xy(w/50+5, 0.7*h)
        pdf_BTech.cell(w=35, h=10, align='L', txt="Date Generated:", border=0)
        pdf_BTech.cell(w=35, h=10, align='C', txt=ind_time, border=0)

        #rectangle
        pdf_BTech.rect( x= 0.2*(w-2*w/50)+w/50, y= w/50+(h-2*(w/50))*0.1+0.016*h, w= 0.6*(w-2*w/50), h= 0.045*h, style='D')

        #lines
        pdf_BTech.line(w/50, 0.63*h+10, w-(w/50), 0.63*h+10) #280.2
        pdf_BTech.line(w/50, w/50+(h-2*(w/50))*0.1, w-(w/50), w/50+(h-2*(w/50))*0.1)
        pdf_BTech.line(0.14*w, w/50+(h-2*(w/50))*0.1, 0.14*w, w/50)
        pdf_BTech.line(w-0.14*w, w/50+(h-2*(w/50))*0.1, w-0.14*w, w/50)

      #rectangles
        pdf_BTech.rect( x= w/50, y= w/50, w= w-2*(w/50), h= h-2*(w/50), style='D')

      #TABLE
        tables_h=w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+10

        data_1=big_data[r]
        if len(data_1)<=8:
          #print(data_1)
          pdf_BTech.set_font("Times", size=7)
          line_height = pdf_BTech.font_size+0.7
          col_width = ((w-2*w/50)-10)/4  # distribute content evenly
          micro_width = []
          micro_width.append(0.15*col_width)
          micro_width.append(0.55*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          pdf_BTech.x=w/50+10/5
          pdf_BTech.y=tables_h
          idx=0
          for row in data_1:
              if(idx<=3):
                #right_max=pdf_BTech.x+micro_width[i-1]+10/5
                i=0
                right_min=pdf_BTech.x
                for items in row:
                    a=right_min
                    right_min=pdf_BTech.x+micro_width[i]
                    for datum in items:
                        #pdf_BTech.x=right_min
                        #pdf_BTech.x=w/50+10/5 
                        pdf_BTech.x=a
                        pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                        
                    pdf_BTech.x=right_min
                    pdf_BTech.y=tables_h
                    i+=1
                pdf_BTech.x=right_min+2
                pdf_BTech.y=tables_h
                idx+=1
              else:
                break
          ht=w/50+10/5

          for i in range(idx):
            pdf_BTech.set_font('Arial', 'BU', 9)
            pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5)
            pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(i+1), border=0)
            rect_text='Credits Taken:  '+str(sem_credits[r][i])+'    Credits Cleared:  '+str(credits_cleared[r][i])+'    SPI:  '+str(spi[r][i])+'    CPI:  '+str(CPI[r][i])
            pdf_BTech.set_xy(ht, 0.35*h)
            pdf_BTech.set_font('Arial', 'B', 8)
            pdf_BTech.cell(w=((w-2*w/50)-10)/4, h=pdf_BTech.font_size*1.2, align='L', txt= rect_text, border=0)
            pdf_BTech.rect(x=ht, y=0.35*h, w=((w-2*w/50)-10)/4, h= pdf_BTech.font_size*1.2)
            ht=ht+col_width+2
          pdf_BTech.line(w/50, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5-7, w-(w/50), w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5-7)
          if(idx>3): #
            pdf_BTech.set_font("Times", size=7)
            line_height = pdf_BTech.font_size+0.7
            col_width = ((w-2*w/50)-10)/4  # distribute content evenly
            #micro_width = col_width/5
            micro_width = []
            micro_width.append(0.15*col_width)
            micro_width.append(0.55*col_width)
            micro_width.append(0.1*col_width)
            micro_width.append(0.1*col_width)
            micro_width.append(0.1*col_width)
            pdf_BTech.x=w/50+10/5
            pdf_BTech.y=tables_h+70
            j=0
            for row in data_1[4:]:
                if j<=3:
                  i=0
                  right_min=pdf_BTech.x
                  for items in row:
                      a=right_min
                      right_min=pdf_BTech.x+micro_width[i]
                      for datum in items:
                          pdf_BTech.x=a
                          pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                          
                      pdf_BTech.x=right_min
                      pdf_BTech.y=tables_h+70
                      i+=1
                  pdf_BTech.x=right_min+2
                  pdf_BTech.y=tables_h+70
                  j+=1
                else:
                  break
            ht=w/50+10/5
            for i in range(j):
              pdf_BTech.set_font('Arial', 'BU', 9)
              pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5)
              pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(idx+i+1), border=0)
              rect_text='Credits Taken:  '+str(sem_credits[r][idx+i])+'    Credits Cleared:  '+str(credits_cleared[r][idx+i])+'    SPI:  '+str(spi[r][idx+i])+'    CPI:  '+str(CPI[r][idx+i])
              pdf_BTech.set_xy(ht, 0.35*h+tables_h+5)
              pdf_BTech.set_font('Arial', 'B', 8)
              pdf_BTech.cell(w=((w-2*w/50)-10)/4, h=pdf_BTech.font_size*1.2, align='L', txt= rect_text, border=0)
              pdf_BTech.rect(x=ht, y=0.35*h+tables_h+5, w=((w-2*w/50)-10)/4, h= pdf_BTech.font_size*1.2)
              ht=ht+col_width+2
        else:
          pdf_BTech.set_font("Times", size=7)
          line_height = pdf_BTech.font_size+0.7
          col_width = ((w-2*w/50)-10)/4  # distribute content evenly
          micro_width = []
          micro_width.append(0.15*col_width)
          micro_width.append(0.55*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          pdf_BTech.x=w/50+10/5
          pdf_BTech.y=tables_h-4.5
          idx=0
          for row in data_1:
              if(idx<=3):
                #right_max=pdf_BTech.x+micro_width[i-1]+10/5
                i=0
                right_min=pdf_BTech.x
                for items in row:
                    a=right_min
                    right_min=pdf_BTech.x+micro_width[i]
                    for datum in items:
                        #pdf_BTech.x=right_min
                        #pdf_BTech.x=w/50+10/5 
                        pdf_BTech.x=a
                        pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                        
                    pdf_BTech.x=right_min
                    pdf_BTech.y=tables_h-4.5
                    i+=1
                pdf_BTech.x=right_min+2
                pdf_BTech.y=tables_h-4.5
                idx+=1
              else:
                break
          ht=w/50+10/5

          for i in range(idx):
            pdf_BTech.set_font('Arial', 'BU', 9)
            pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5-4)
            pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(i+1), border=0)
            #Credits Rectangle
            rect_text='Credits Taken:  '+str(sem_credits[r][i])+"    Credits Cleared:  "+str(credits_cleared[r][i])+"    SPI:  "+str(spi[r][i])+"    CPI:  "+str(CPI[r][i])
            pdf_BTech.set_font('Arial', 'B', 7)
            creds_y=tables_h+9+pdf_BTech.font_size*12
            pdf_BTech.set_xy(ht, creds_y-7)
            pdf_BTech.cell(w=((w-2*w/50)-10)/4, h=pdf_BTech.font_size*1.3, align='L', txt= rect_text, border=1)
            #pdf_BTech.rect(x=ht, y=0.35*h, w=((w-2*w/50)-10)/4, h= pdf_BTech.font_size*1.2)
            ht=ht+col_width+2
          pdf_BTech.line(w/50, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5-10-15+2-3, w-w/50, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5-10-15+2-3)
          if(idx>3): #
            pdf_BTech.set_font("Times", size=7)
            line_height = pdf_BTech.font_size+0.7
            col_width = ((w-2*w/50)-10)/4  # distribute content evenly
            #micro_width = col_width/5
            micro_width = []
            micro_width.append(0.15*col_width)
            micro_width.append(0.55*col_width)
            micro_width.append(0.1*col_width)
            micro_width.append(0.1*col_width)
            micro_width.append(0.1*col_width)
            pdf_BTech.x=w/50+10/5
            pdf_BTech.y=tables_h+70-22.5
            j=0
            for row in data_1[4:]:
              if j<=3:
                i=0
                right_min=pdf_BTech.x
                for items in row:
                    a=right_min
                    right_min=pdf_BTech.x+micro_width[i]
                    for datum in items:
                        pdf_BTech.x=a
                        pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                        
                    pdf_BTech.x=right_min
                    pdf_BTech.y=tables_h-4.5+0.175*h
                    i+=1
                pdf_BTech.x=right_min+2
                pdf_BTech.y=tables_h-4.5+0.175*h
                j+=1
              else:
                break
            ht=w/50+10/5
            for i in range(j):
              pdf_BTech.set_font('Arial', 'BU', 9)
              pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5-10-15+2)
              pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(idx+i+1), border=0)
              #Credits Rectangle
              rect_text='Credits Taken:  '+str(sem_credits[r][idx+i])+'    Credits Cleared:  '+str(credits_cleared[r][idx+i])+'    SPI:  '+str(spi[r][idx+i])+'    CPI:  '+str(CPI[r][idx+i])
              pdf_BTech.set_font('Arial', 'B', 7)
              creds_y=tables_h+9+pdf_BTech.font_size*12
              pdf_BTech.set_xy(ht, creds_y+17+15+6)
              pdf_BTech.cell(w=((w-2*w/50)-10)/4, h=pdf_BTech.font_size*1.3, align='L', txt= rect_text, border=1)
              ht=ht+col_width+2
            pdf_BTech.line(w/50, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5+8+7+13-3, w-w/50, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5+8+7+13-3)

            if(j>3): #
              pdf_BTech.set_font("Times", size=7)
              line_height = pdf_BTech.font_size+0.7
              col_width = ((w-2*w/50)-10)/4  # distribute content evenly
              #micro_width = col_width/5
              micro_width = []
              micro_width.append(0.15*col_width)
              micro_width.append(0.55*col_width)
              micro_width.append(0.1*col_width)
              micro_width.append(0.1*col_width)
              micro_width.append(0.1*col_width)
              pdf_BTech.x=w/50+10/5
              pdf_BTech.y=tables_h-4.5+2*0.175*h
              kk=0
              for row in data_1[8:]:
                if kk<=3:
                  i=0
                  right_min=pdf_BTech.x
                  for items in row:
                      a=right_min
                      right_min=pdf_BTech.x+micro_width[i]
                      for datum in items:
                          pdf_BTech.x=a
                          pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                          
                      pdf_BTech.x=right_min
                      pdf_BTech.y=tables_h-4.5+2*0.175*h
                      i+=1
                  pdf_BTech.x=right_min+2
                  pdf_BTech.y=tables_h-4.5+2*0.175*h
                  kk+=1
              ht=w/50+10/5
              for i in range(kk):
                pdf_BTech.set_font('Arial', 'BU', 9)
                pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5+8+7+13)
                pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(idx+j+i+1), border=0)
                #Credits Rectangle
                rect_text='Credits Taken:  '+str(sem_credits[r][idx+i+j])+'    Credits Cleared:  '+str(credits_cleared[r][idx+i+j])+'    SPI:  '+str(spi[r][idx+i+j])+'    CPI:  '+str(CPI[r][idx+i+j])
                pdf_BTech.set_font('Arial', 'B', 7)
                creds_y=tables_h+9+pdf_BTech.font_size*12
                pdf_BTech.set_xy(ht, creds_y+30+25+18)
                pdf_BTech.cell(w=((w-2*w/50)-10)/4, h=pdf_BTech.font_size*1.3, align='L', txt= rect_text, border=1)
                ht=ht+col_width+2

      else:
        #pdf_Rest = PDF(orientation='P', unit='mm', format='A4')
        pdf_BTech = PDF(orientation='P', unit='mm', format='A4')
        w=210
        h=297
        #Images
        pdf_BTech.set_xy(0.12*w/6+w/50, w/50+(h-2*(w/50))*0.1*0.1)
        pdf_BTech.image('iitp-1_black.png',  link='', type='', w=0.08*w, h=(h-2*(w/50))*0.1*0.7)
        pdf_BTech.set_xy(w-0.14*w+0.12*w/6, w/50+(h-2*(w/50))*0.1*0.1)
        pdf_BTech.image('iitp-1_black.png',  link='', type='', w=0.08*w, h=(h-2*(w/50))*0.1*0.7)
        pdf_BTech.set_xy(0.14*w, w/50)
        pdf_BTech.image('iitp_heading.png',  link='', type='', w=0.67*w, h=0.081*h)
        if stamp is not None:
          pdf_BTech.set_xy(0.38*w, 0.7*h)
          pdf_BTech.image('stamp_iitp.png',  link='', type='', w=40.03/2, h=40.48/2)
        if sign is not None:
          pdf_BTech.set_xy(0.825*w, 0.725*h)
          pdf_BTech.image('assistant_reg.png',  link='', type='', w=40.03/2, h=21.08/2)

        #Text
        pdf_BTech.set_font('Arial', 'B', 6)
        pdf_BTech.set_text_color(0, 0, 0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Roll No.:", border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h+0.035*h/2)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Programme:", border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Name:", border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h+0.035*h/2)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Course:", border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+2*0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt="Year of Admission:", border=0)


        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=str(r), border=0)

        prog={"11":"Master of Technology","12":"Master of Science","21":"Doctor of Philosophy"}
    
        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5, w/50+(h-2*(w/50))*0.1+0.016*h+0.035*h/2)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=prog[r[2:4]], border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=roll_names[r], border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h+0.035*h/2)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt=str(r[4:6]), border=0)

        pdf_BTech.set_xy(0.15*(w-2*w/50)+w/50+5+2*0.7*(w-2*w/50)/3, w/50+(h-2*(w/50))*0.1+0.016*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='C', txt="Year of Admission: "+year, border=0)

        pdf_BTech.set_font('Arial', 'B', 5)
        pdf_BTech.set_xy(0.8*w, 0.85*h)
        pdf_BTech.cell(w=0.7*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Assistant Registrar (Academic)", border=0)


        #rectangle
        pdf_BTech.rect( x= 0.15*(w-2*w/50)+w/50, y= w/50+(h-2*(w/50))*0.1+0.016*h, w= 0.7*(w-2*w/50), h= 0.035*h, style='D')

        #lines
        pdf_BTech.line(w/50, w/50+(h-2*(w/50))*0.1, w-(w/50), w/50+(h-2*(w/50))*0.1) #
        pdf_BTech.line(w/50, 0.35*h, w-(w/50), 0.35*h) #
        pdf_BTech.line(w/50, 0.55*h, w-(w/50), 0.55*h) #
        pdf_BTech.line(w/50, 0.75*h, w-(w/50), 0.75*h) #
        pdf_BTech.line(0.14*w, w/50+(h-2*(w/50))*0.1, 0.14*w, w/50) #
        pdf_BTech.line(w-0.14*w, w/50+(h-2*(w/50))*0.1, w-0.14*w, w/50) #

        #main rectangle
        pdf_BTech.rect( x= w/50, y= w/50, w= w-2*(w/50), h= h-2*(w/50), style='D')

        tables_h=w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+10

        data_1=big_data[r]
        #print(data_1)
        pdf_BTech.set_font("Times", size=7)
        line_height = pdf_BTech.font_size+1.1
        col_width = ((w-2*w/50)-10)/2  # distribute content evenly
        micro_width = []
        micro_width.append(0.15*col_width)
        micro_width.append(0.55*col_width)
        micro_width.append(0.1*col_width)
        micro_width.append(0.1*col_width)
        micro_width.append(0.1*col_width)
        pdf_BTech.x=w/50+10/5
        pdf_BTech.y=tables_h-4.5
        idx=0
        for row in data_1:
            if(idx<=1):
              #right_max=pdf_BTech.x+micro_width[i-1]+10/5
              i=0
              right_min=pdf_BTech.x
              for items in row:
                  a=right_min
                  right_min=pdf_BTech.x+micro_width[i]
                  for datum in items:
                      #pdf_BTech.x=right_min
                      #pdf_BTech.x=w/50+10/5 
                      pdf_BTech.x=a
                      pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                      
                  pdf_BTech.x=right_min
                  pdf_BTech.y=tables_h-4.5
                  i+=1
              pdf_BTech.x=right_min+2
              pdf_BTech.y=tables_h-4.5
              idx+=1
            else:
              break
        ht=w/50+10/5

        for i in range(idx):
          pdf_BTech.set_font('Arial', 'BU', 9)
          pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5)
          pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(i+1), border=0)
          #Credits Rectangle
          rect_text='Credits Taken:  '+str(sem_credits[r][i])+"    Credits Cleared:  "+str(credits_cleared[r][i])+"    SPI:  "+str(spi[r][i])+"    CPI:  "+str(CPI[r][i])
          pdf_BTech.set_font('Arial', 'B', 7)
          creds_y=tables_h+9+pdf_BTech.font_size*12
          pdf_BTech.set_xy(ht, creds_y)
          pdf_BTech.cell(w=((w-2*w/50)-10)/2, h=pdf_BTech.font_size*1.2, align='L', txt= rect_text, border=1)
          #pdf_BTech.rect(x=ht, y=0.35*h, w=((w-2*w/50)-10)/4, h= pdf_BTech.font_size*1.2)
          ht=ht+col_width+2
        if(idx>1): #
          pdf_BTech.set_font("Times", size=7)
          line_height = pdf_BTech.font_size+1.1
          col_width = ((w-2*w/50)-10)/2  # distribute content evenly
          #micro_width = col_width/5
          micro_width = []
          micro_width.append(0.15*col_width)
          micro_width.append(0.55*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          pdf_BTech.x=w/50+10/5
          pdf_BTech.y=tables_h
          j=0
          for row in data_1[2:]:
            if j<=1:
              i=0
              right_min=pdf_BTech.x
              for items in row:
                  a=right_min
                  right_min=pdf_BTech.x+micro_width[i]
                  for datum in items:
                      pdf_BTech.x=a
                      pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                      
                  pdf_BTech.x=right_min
                  pdf_BTech.y=tables_h-4.5+0.175*h
                  i+=1
              pdf_BTech.x=right_min+2
              pdf_BTech.y=tables_h-4.5+0.175*h
              j+=1
            else:
              break
          ht=w/50+10/5
          for i in range(j):
            pdf_BTech.set_font('Arial', 'BU', 9)
            pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5)
            pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(idx+i+1), border=0)
            #Credits Rectangle
            rect_text='Credits Taken:  '+str(sem_credits[r][idx+i])+'    Credits Cleared:  '+str(credits_cleared[r][idx+i])+'    SPI:  '+str(spi[r][idx+i])+'    CPI:  '+str(CPI[r][idx+i])
            pdf_BTech.set_font('Arial', 'B', 7)
            creds_y=tables_h+9+pdf_BTech.font_size*12
            pdf_BTech.set_xy(ht, creds_y)
            pdf_BTech.cell(w=((w-2*w/50)-10)/2, h=pdf_BTech.font_size*1.2, align='L', txt= rect_text, border=1)
            ht=ht+col_width+2

        if(j>1): #
          pdf_BTech.set_font("Times", size=7)
          line_height = pdf_BTech.font_size+2
          col_width = ((w-2*w/50)-10)/2  # distribute content evenly
          #micro_width = col_width/5
          micro_width = []
          micro_width.append(0.15*col_width)
          micro_width.append(0.55*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          micro_width.append(0.1*col_width)
          pdf_BTech.x=w/50+10/5
          pdf_BTech.y=tables_h-4.5+2*0.175*h
          kk=0
          for row in data_1[4:]:
              i=0
              right_min=pdf_BTech.x
              for items in row:
                  a=right_min
                  right_min=pdf_BTech.x+micro_width[i]
                  for datum in items:
                      pdf_BTech.x=a
                      pdf_BTech.multi_cell(micro_width[i], line_height, str(datum), border=1, align='C')
                      
                  pdf_BTech.x=right_min
                  pdf_BTech.y=tables_h-4.5+2*0.175*h
                  i+=1
              pdf_BTech.x=right_min+2
              pdf_BTech.y=tables_h-4.5+2*0.175*h
              kk+=1
          ht=w/50+10/5
          for i in range(kk):
            pdf_BTech.set_font('Arial', 'BU', 9)
            pdf_BTech.set_xy(ht, w/50+(h-2*(w/50))*0.1+0.016*h+0.045*h+4.5+tables_h+5)
            pdf_BTech.cell(w=0.6*(w-2*w/50)/3, h=0.045*h/2, align='L', txt="Semester "+str(idx+i+1), border=0)
            #Credits Rectangle
            rect_text='Credits Taken:  '+str(sem_credits[r][idx+i+j])+'    Credits Cleared:  '+str(credits_cleared[r][idx+i+j])+'    SPI:  '+str(spi[r][idx+i+j])+'    CPI:  '+str(CPI[r][idx+i+j])
            pdf_BTech.set_font('Arial', 'B', 7)
            creds_y=tables_h+9+pdf_BTech.font_size*12
            pdf_BTech.set_xy(ht, creds_y)
            pdf_BTech.cell(w=((w-2*w/50)-10)/2, h=pdf_BTech.font_size*1.2, align='L', txt= rect_text, border=1)
            ht=ht+col_width+2
      file_name=path+"/"+r+".pdf"
      pdf_BTech.output(name= file_name, dest='F')
    else:
      invalid_rolls.append(r)
  return invalid_rolls