import openpyxl
from io import BytesIO
import streamlit as st

st.title("Language and Cognition Lab🧠📊")
st.image('https://i.ibb.co/L0hHCFZ/Screenshot-2023-02-22-at-12-57-43-PM.png')

st.markdown("Automated Excel Editor")
excel_file = st.file_uploader('Upload your excel file')

def pre_process_excel():
    wb = openpyxl.load_workbook(BytesIO(excel_file.read()))
    sheet = wb.active #specify the sheet name to select other than the active sheet
    print("Maximum rows before removing : ",sheet.max_row)
    sheet.insert_cols(1, 2)
    sheet.insert_cols(idx=4)
    sheet.delete_rows(0,17)
    sheet['A1'] = "Subject Number"
    sheet['B1'] = 'list'
    sheet['D1'] = 'null'
    for row in sheet.iter_rows():
        values = [cell.value for cell in row]
        values[4], values[10] = values[10], values[4]
        for i, cell in enumerate(row):
            cell.value = values[i]
    # wb.save()
    return sheet


def edit_excel():
  path1 = 'edited_.xlsx'
  wb = openpyxl.Workbook()
  ws_old = pre_process_excel()
  column_values = ['Subject Number', 'list', 'trial', 'null', 'condition', 'time',
                   'Relationship',
                   'ControlQ1 Copy 2', 'ControlQ1 Copy - 2 - 2',
                   'FirstMoozleProp Copy 13',
                  'SecondMoozleProp Copy 13', 'SecondMoozleProp2 Copy 13','ChoiceResponse Copy 2',
                  'ControlQ2 Copy 2', 'ControlQ2 Copy-2 - 2', 'Choice','SameChoice ', 'BeliefType', 'AgeGroup']
  ws = wb.active


  for index, value in enumerate(column_values, start=1):
      cell = ws.cell(row=1, column=index)
      cell.value = value
      wb.save(path1)

  i = 2 # sample numbers start at 2
  j = 11 # Every sample has 9 trials, hence the hence sample is i + 9 = 11
  while i <= 25:
    for cell in ws_old['C']:
        if cell.value == i:
            if i == 2:
              # Trial
              string = ws_old.cell(row=i, column=3).value
              ws.cell(row=i, column=3).value = string

               # Condition
              string = ws_old.cell(row=i, column=6).value
              ws.cell(row=i, column=5).value = string.replace('.PICT @ :Pictures:', '')

              # Relationship
              string = ws_old.cell(row=i, column=5).value
              ws.cell(row=i, column=7).value = string.replace('.PICT @ :Pictures:', '')
              
              #ControlQ1 Copy 2 the 2nd row for the trial in keys between []
              string = ws_old.cell(row=i+1, column=14).value
              ws.cell(row=i, column=8).value = string.replace('[', '').replace(']','')

              # ControlQ1 Copy - 2 - 2the 3rd  row for the trial in keys
              string = ws_old.cell(row=i+2, column=14).value
              ws.cell(row=i, column=9).value = string.replace('[', '').replace(']','')
 
              #FirstMoozleProp Copy 13 the 4th row for the trial in response label
              string = ws_old.cell(row=i+3, column=5).value
              ws.cell(row=i, column=10).value = string.replace('.PICT @ :Pictures:', '')

              # SecondMoozleProp Copy 13 the 5th row for the trial in response label
              string = ws_old.cell(row=i+4, column=5).value
              ws.cell(row=i, column=11).value = string.replace('.PICT @ :Pictures:', '')


              #SecondMoozleProp2 Copy 13he 6th row for the trial in response
              string = ws_old.cell(row=i+5, column=5).value
              ws.cell(row=i, column=12).value = string.replace('.PICT @ :Pictures:', '')

              # time the 6th row for that trial in Time
              string = ws_old.cell(row=i+6, column=12).value
              ws.cell(row=i, column=6).value = string

              #ChoiceResponse Copy 2 the 7th row for the trial in keys between []
              string = ws_old.cell(row=i+6, column=14).value
              ws.cell(row=i, column=13).value = string.replace('[', '').replace(']','')


              #ControlQ2 Copy 2 the 8th row for the trial in keys between
              string = ws_old.cell(row=i+7, column=14).value
              ws.cell(row=i, column=14).value = string.replace('[', '').replace(']','')
              
              #ControlQ2 Copy-2 - 2 the 9th row for the trial in keys between
              string = ws_old.cell(row=i+8, column=14).value
              ws.cell(row=i, column=15).value = string.replace('[', '').replace(']','')

              i += 1
            else:
              # Trial
              string = ws_old.cell(row=j, column=3).value
              ws.cell(row=i, column=3).value = string

              # Condition
              string = ws_old.cell(row=j, column=6).value
              ws.cell(row=i, column=5).value = string.replace('.PICT @ :Pictures:', '')

              # Relationship
              string = ws_old.cell(row=j, column=5).value
              ws.cell(row=i, column=7).value = string.replace('.PICT @ :Pictures:', '')

              #ControlQ1 Copy 2 the 2nd row for the trial in keys between []
              string = ws_old.cell(row=j+1, column=14).value
              ws.cell(row=i, column=8).value = string.replace('[', '').replace(']','')

 
              #ControlQ1 Copy - 2 - 2the 3rd  row for the trial in keys
              string = ws_old.cell(row=j+2, column=14).value
              ws.cell(row=i, column=9).value = string.replace('[', '').replace(']','')
 

              #FirstMoozleProp Copy 13 the 4th row for the trial in response label
              string = ws_old.cell(row=j+3, column=5).value
              ws.cell(row=i, column=10).value = string.replace('.PICT @ :Pictures:', '')

              # SecondMoozleProp Copy 13 the 5th row for the trial in response label
              string = ws_old.cell(row=j+4, column=5).value
              ws.cell(row=i, column=11).value = string.replace('.PICT @ :Pictures:', '')

              #SecondMoozleProp2 Copy 13he 6th row for the trial in response
              string = ws_old.cell(row=j+5, column=5).value
              ws.cell(row=i, column=12).value = string.replace('.PICT @ :Pictures:', '')


              # time the 6th row for that trial in Time
              string = ws_old.cell(row=j+6, column=12).value
              ws.cell(row=i, column=6).value = string

              #ChoiceResponse Copy 2 the 7th row for the trial in keys between []
              string = ws_old.cell(row=j+6, column=14).value
              ws.cell(row=i, column=13).value = string.replace('[', '').replace(']','')

              #ControlQ2 Copy 2 the 8th row for the trial in keys between
              string = ws_old.cell(row=j+7, column=14).value
              ws.cell(row=i, column=14).value = string.replace('[', '').replace(']','')

              #ControlQ2 Copy-2 - 2 the 9th row for the trial in keys between
              string = ws_old.cell(row=j+8, column=14).value
              ws.cell(row=i, column=15).value = string.replace('[', '').replace(']','')

              i += 1
              j += 9  
  path1 = 'edited_.xlsx'
  wb.save(path1)
  st.download_button(
        label="Download Updated Excel Workbook",
        data=open(path1, 'rb').read(),
        file_name="workbook.xlsx",
        mime="xlsx"
    )



  
      
edit = st.button('Click me')
if edit:
    edit_excel()