import streamlit as st
import pandas as pd
import base64
from io import BytesIO

def to_excel(dfp):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    dfp.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def load(dfp_kd,dfp_fz):
    path_kd_use_dict = {'运单号':'快递单号',
                 '订单编号':'编号',
                 '收件人姓名':'收件人',
                 '收件人手机':'电话',}
    path_fz_str_cols = ['电话','发件方号码']
    dfp_kd = dfp_kd[path_kd_use_dict.keys()]\
        .rename(columns = path_kd_use_dict)\
        .astype('str')
    #去掉重复行
    dfp_kd = dfp_kd.drop_duplicates(keep='first')

    dfp_fz = dfp_fz
    if '快递单号' in dfp_fz.columns:
        dfp_fz = dfp_fz.drop(columns=['快递单号'])
    dfp_fz[path_fz_str_cols] = dfp_fz[path_fz_str_cols].astype('str')


    return dfp_kd,dfp_fz

def join_data(dfp_kd,dfp_fz):
    return dfp_fz.merge(dfp_kd,on=['编号','收件人','电话'],how='left')

def run(path_kd,path_fz,save_path):
    dfp_kd, dfp_fz = load(path_kd, path_fz)
    dfp_out = join_data(dfp_kd, dfp_fz)
    dfp_out.to_excel(save_path, index=False)

    return 1

if __name__ == "__main__":
    uploaded_file_kd = st.file_uploader("上传kd文件", type="xls")
    uploaded_file_fz = st.file_uploader("上传fz文件", type="xlsx")
    if uploaded_file_kd is not None and uploaded_file_fz is not None:
        dfp_kd = pd.read_excel(uploaded_file_kd)
        dfp_fz = pd.read_excel(uploaded_file_fz)

        dfp_kd, dfp_fz = load(dfp_kd,dfp_fz)
        dfp_out = join_data(dfp_kd, dfp_fz)
        st.write('kd 展示')
        st.write(dfp_kd)
        st.write('fz 展示')
        st.write(dfp_fz)
        st.write('out 展示')
        st.write(dfp_out)

        val =  to_excel(dfp_out)
        b65 = base64.b64encode(val)
        href = f'<a href="data:application/octet-stream;base64,{b65.decode()}" download="extract.xlsx">Download excel file</a>' # decode b'abc' => abc
        st.markdown(href, unsafe_allow_html=True)


