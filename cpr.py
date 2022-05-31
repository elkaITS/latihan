from csv import excel
import streamlit as st
import pandas as pd
import re
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

# Header
st.title("Coordinate Converter")
st.subheader("Selamat Datang di Website Aplikasi Coordinate Converter")
st.caption("Website ini berisi aplikasi yang mampu digunakan dalam memudahkan proyek sehubungan dengan penentuan titik koordinat.")
st.image("https://conocedores.com/wp-content/uploads/2016/06/googlemapslogo.jpg")


# Input User
with st.sidebar:
    st.subheader("Menu")
    p = st.selectbox('Pilih Fungsi yang Anda Butuhkan',('Home','Coordinate Converter','Input Data'))
    


# Page Home
if p == 'Home':
    st.snow()
# Page Converter
if p == 'Coordinate Converter':
    r = st.selectbox('Pilih Besaran yang ingin anda cari',('None','Latitude','Longitude'))
    if r == 'None':
        st.snow()
    if r == 'Latitude':
        a = st.text_input('Masukan Latitude', value="°\'\"", max_chars=14, )
        deg, min, s, dr = re.split('[°\'\"]', a)
        g = st.button("Lihat Hasil")
        if g:
            m = float(deg)
            n = float(min)
            o = float(s)
            rumus1 = (m+(n/60)+(o/3600))*(-1)
            st.success(f"Hasilnya adalah {rumus1}")
        
    if r == 'Longitude':
        b = st.text_input('Masukan Longitude', value="°\'\"", max_chars=14, )
        gree, mnt, sec, dir = re.split('[°\'\"]', b)
        h = st.button("Lihat Hasil")
        if h:
            u = float(gree)
            v = float(mnt)
            w = float(sec)
            rumus2 = (u+(v/60)+(w/3600))
            st.success(f"Hasilnya adalah {rumus2}")


# Input Manual

if p == 'Input Data':
    q = st.file_uploader('Pilih File yang ingin diupload', type='xlsx')
    if q:
        st.markdown('---')
        st.subheader('Data Anda telah Terupload!')
        df=pd.read_excel(q, engine='openpyxl')
        st.dataframe(df)
        t=len(df)
        La=df['Latitude'].tolist()
        row1 = []
        for i in La:
            deg, min, s, dr = re.split('[°\'\"]', i)
            k = float(deg)
            l = float(min)
            q = float(s)
            rumus3 = (k+(l/60)+(q/3600))*(-1)
            ku = row1.append([i,rumus3])
        data_1 = pd.DataFrame(row1,columns=['Latitude','Converted Latitude'])
        Lo = df['Longitude'].to_list()
        row2 = []
        for i in Lo:
            der,men,det,arah = re.split('[°\'\"]', i)
            de = float(deg)
            me = float(min)
            det = float(s)
            rumus4 = (me+(de/60)+(det/3600))*(-1)
            kj = row2.append([i,rumus4])
        data_2 = pd.DataFrame(row2,columns=['Longitude','Converted Longitude'])
        st.write("Maka hasilnya adalah")
        data_fix = data_1.join(data_2)
        data_download = data_fix.drop(columns=['Latitude','Longitude'])
        st.dataframe(data_fix)
            
        # Export Excel
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            format1 = workbook.add_format({'num_format': '0.00'}) 
            worksheet.set_column('A:A', None, format1)  
            writer.save()
            processed_data = output.getvalue()
            return processed_data
        df_xlsx = to_excel(data_download)
        st.download_button(label='📥 Download Hasil Ke Excel',
                                        data=df_xlsx ,
                                        file_name= 'df_test.xlsx')


