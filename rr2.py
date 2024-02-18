# -*- coding: utf-8 -*-
"""
Created on Fri Jan 26 08:36:25 2024

@author: zijie.xu
"""

import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw
import tkinter

import pickle
from PIL import Image
from tkinter import filedialog
st.title('Tire RR Values Prediction')
st.write(':blue[Only] :blue[application] :blue[to] :red[R16,R17,R18] :blue[Tires]')
rrmodelfile ='static/1.dat'
#rrmodelfile = '\\10.97.1.43\赛轮集团股份有限公司\技术研发中心\研发实验中心\技术研发中心实验系统\1.中心实验室\1.中心实验室\9-实验中心-成品组\03 六分力组\5、数据处理\rr\R16_R17_R18_M8(104,10).dat'
st.write(rrmodelfile)
load_model=pickle.load(open(rrmodelfile))

rr0=0
with st.container():
    st.write('Single RR Value Prediction:')
    column=['Groove Depth','Tread Surface Width','TR1','TR2','Tread Tan D','Tread Area','Base Area','Sidewall Area','Sidewall Tan D','Belt Width1','Belt Density','Belt Angle','Carcass Material','Turn Up1','Turn Up2','Apex Height','Apex Tan D','Rim Width','Infation Pressure','Load','Size']
    df=pd.DataFrame(np.ones((1,21)),columns=column)
    dfnew=st.data_editor(df)
    col1, col2 = st.columns(2)
    with col1:
        st.button('RR Prediction:')
    with col2:
        dfnew=np.array(dfnew)
        rr0 = load_model.predict(dfnew)/dfnew[0,-2]*1000
        st.markdown(':red[RR] :red[Value:] :red[%.3f]'%rr0)
    st.write('RR Excel File Prediction:')
    col3, col4 = st.columns(2)
    with col3:
        if st.button('File Read'):
            root = tkinter.Tk()
            root.withdraw()
            
            top = tkinter.Toplevel(root)
            top.title("选择文件")
            f_path = filedialog.askopenfilename(parent=top)
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(f_path) # 打开Excel文件
            sheet = wb.sheets[0]  # 选择第0个表单
            list_value = sheet.range('A1').expand().value
            df00=pd.DataFrame(list_value[1:],columns=list_value[0])
            df00=np.array(df00)
          #  st.write(df00[0,:])
            rr_file = load_model.predict(df00)/df00[:,-2]*1000
            x_new=np.array(rr_file).reshape(len(rr_file),1)
            ynew = np.append(df00,x_new,axis=1)
            column1=['Groove Depth','Tread Surface Width','TR1','TR2','Tread Tan D','Tread Area','Base Area','Sidewall Area','Sidewall Tan D','Belt Width1','Belt Density','Belt Angle','Carcass Material','Turn Up1','Turn Up2','Apex Height','Apex Tan D','Rim Width','Infation Pressure','Load','Size','RR Prediction']

            ynew = pd.DataFrame(ynew,columns=column1)
        else:
            ynew = pd.DataFrame(np.ones((1,1)),columns=['0'])
            st.write('文件未读取！')
    with col4:
        #@st.cache
        def convert_df(df):
            # IMPORTANT: Cache the conversion to prevent computation on every rerun
            return df.to_csv().encode('utf-8')
        csv = convert_df(ynew)
        
        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name='RR_Predict_Data.csv',
            mime='text/csv',)
    image = Image.open(r'\\10.97.1.43\赛轮集团股份有限公司\技术研发中心\研发实验中心\技术研发中心实验系统\1.中心实验室\1.中心实验室\9-实验中心-成品组\03 六分力组\5、数据处理\rr\rsq.jpg')
    st.image(image)
    st.write('Predict Model RSQ:0.9901')
    st.write('Model Update: 2024.01.26')


    

    
