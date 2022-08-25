from re import L
import pandas as pd


import os
import xlsxwriter


campos=['id','name','Company','Related Company /External ID']
excelContactos=pd.read_excel('ContactoYCompany.xlsx',sheet_name='Sheet1')
excelCompanies=pd.read_excel('ContactoYCompany.xlsx',sheet_name='Sheet2')
excelContactos=excelContactos.fillna(0)
listaContactos=excelContactos.values.tolist()
excelCompanies=excelCompanies.fillna(0)
listaCompanies=excelCompanies.values.tolist()
print(listaContactos)
print(listaCompanies)
with xlsxwriter.Workbook('Asosiacion.xlsx' ) as workbook:
            worksheet=workbook.add_worksheet()
            row = 0
            col = 0
 
        # Iterate over the data and write it out row by row.
            for campo in campos:
                 bold = workbook.add_format({'bold': True})
                 worksheet.write( row,col, campo,bold)  
                 col += 1
            col=0  
            
            row+=1
            for  elementos in listaContactos:
               
    #e=listaContactos[index+2]
 #for  elem in elementos:
                worksheet.write(row, 0, elementos[0])
                col+=1
                worksheet.write(row,1, elementos[1])
                col+=1
                worksheet.write(row, 2, elementos[2])
                col+=1
                
                for  compa in listaCompanies:
                     if elementos[2] in compa:
                       worksheet.write(row, 3,compa[1] )
                       col+=1
                       
                       
                     else:
                      col+=1
                      
                      worksheet.write(row,col, " ") 
                row+=1
                workbook.close
pass
pass
pass