import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
from xlsxwriter.utility import xl_rowcol_to_cell
def to_excel1(df):
    df = df[[df.columns[2],df.columns[1],df.columns[4]]]
    df = df.rename(columns={df.columns[0] : 'NAMA LENGKAP', df.columns[1]: 'KELAS',df.columns[2]: 'NILAI'})
    df['NAMA LENGKAP']= df['NAMA LENGKAP'].str.upper().str.title()
    
    df = df.sort_values(['NAMA LENGKAP'], ascending=[True])
    df1 = df.groupby('KELAS').agg({"count"})
    df1 = df1.reset_index()

    name_sheet = df1["KELAS"].values.tolist()
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    for sheet in name_sheet:
       df[df["KELAS"]== str(sheet)].to_excel(writer, sheet_name=str(sheet),index=False,startcol=0,startrow=4)
    workbook = writer.book
    header_format = workbook.add_format({
            "valign": "vcenter",
            "align": "center",
            "bold": True,
        })
    col_format = workbook.add_format({
            "valign": "vcenter",
            "align": "center",
        })
    col1_format = workbook.add_format()
    for sheet in name_sheet:
      #add title
      title_new = title1
      subheader_new = subheader
      subheader1_new = subheader1
      #merge cells

      format = workbook.add_format({
          "valign": "vcenter",
          "align": "center",
          "bold": True,
      })
      format.set_font_size(12)
      format.set_font_name('Arial')
      header_format.set_font_size(12)
      header_format.set_font_name('Arial')
      col_format.set_font_size(12)
      col_format.set_font_name('Arial')
      col1_format.set_font_size(12)
      col1_format.set_font_name('Arial')
      writer.sheets[sheet].merge_range('A1:C1', title1, format)
      writer.sheets[sheet].merge_range('A2:C2', subheader_new,format)
      writer.sheets[sheet].merge_range('A3:C3', subheader1_new,format)
      writer.sheets[sheet].set_row(2, 15) # Set the header row height to 15
      for col_num, value in enumerate(df[["NAMA LENGKAP","KELAS","NILAI"]].columns.values):
          writer.sheets[sheet].write(4, col_num,value,header_format)
          # Adjust the column width.
          writer.sheets[sheet].set_column('A:A', 40,col1_format)
          writer.sheets[sheet].set_column('B:D', 15,col_format)
      writer.sheets[sheet].conditional_format(xlsxwriter.utility.xl_range(4, 0, 4+len(df[df["KELAS"]== str(sheet)]), len(df[df["KELAS"]== str(sheet)].columns) - 1), {'type': 'no_errors'})
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.write('# REKAP NILAI')
title1 =st.text_input('Judul', 'REKAPITULASI')

subheader =st.text_input('SubJudul', 'PENDIDIKAN AGAMA ISLAM')
subheader1 =st.text_input('SubJudul', 'PTS GANJIL 2022/2023')

uploaded_file = st.file_uploader("Upload spreadsheet", type=["csv", "xlsx"])

# Check if file was uploaded
st.write('Judul : ', title1)
st.write('Sub Judul : ', subheader)
if uploaded_file:
    # Check MIME type of the uploaded file
    if uploaded_file.type == "text/csv":
        df = pd.read_csv(uploaded_file,sep = ';')
    else:
        df = pd.read_excel(uploaded_file)
    df_xlsx = to_excel1(df)
    st.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'NILAI '+subheader+'.xlsx')





