import openpyxl
from io import BytesIO

def edit_excel():
  path1 = 'edited_.xlsx'
  wb = openpyxl.Workbook()
  wb_old = openpyxl.load_workbook('/content/workbook (5).xlsx')
  column_values = ['Subject Number', 'list', 'trial', 'null', 'condition', 'time',
                   'Relationship',
                   'ControlQ1 Copy 2', 'ControlQ1 Copy - 2 - 2',
                   'FirstMoozleProp Copy 13',
                  'SecondMoozleProp Copy 13', 'SecondMoozleProp2 Copy 13','ChoiceResponse Copy 2',
                  'ControlQ2 Copy 2', 'ControlQ2 Copy-2 - 2', 'Choice','SameChoice ', 'BeliefType', 'AgeGroup']
  ws = wb.active
  ws_old = wb_old.active
  ws.delete_rows(0,17)

  for index, value in enumerate(column_values, start=1):
      cell = ws.cell(row=1, column=index)
      cell.value = value
      wb.save(path1)

  # for row in ws.iter_rows():
  #   values = [cell.value for cell in row]
  #   for i, cell in enumerate(row):
  #       cell.value = values[i]    

  # ws.insert_cols(1, 2)
  # ws_old.delete_rows(0,17)
  i = 2
  j = 11
  while i <= 25:
    for cell in ws_old['C']:
        if cell.value == i:
            if i == 2:
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



              i += 1
            else:
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


              i += 1
              j += 9  
  path1 = 'edited_.xlsx'
  wb.save(path1)



  
      

edit_excel()