from numpy import row_stack
import openpyxl
import math

# userInput = input('Ingrese Usuario: ')
# passwordInput = input('Ingrese Contraseña: ')

# usersDB = [
#   {
#     "user": 'Pepito',
#     "password": '123'
#   },
#   {
#     "user": 'Juan',
#     "password": '158'
#   },
#   {
#     "user": 'Pepe',
#     "password": '948'
#   },
# ]

# userExist = False
# for userDB in usersDB:
#   if (userDB['user'] == userInput):
#     userExist = True
#     if (userDB['password'] == passwordInput):
#       print('-- Logeado --')
#       break
#     else:
#       print('-- Contraseña Incorrecta --')
#       break
  
# if (not userExist):
#   print(f'-- El Usuario "{userInput}" No Existe --')

# Init
wb = openpyxl.load_workbook('./Excel.xlsx')
hoja = 'Transito'

if (hoja == 'Hoja1'):
  print('Hoja1')
  sheet = wb['Hoja1']
  rows = sheet.max_row
  columns = sheet.max_column

  count = 1
  rowsExcel = []

  # FOR's para guardar los datos del Excel en un array
  for row in range(2, rows-1):
    rowArr = []
    for column in range(columns):
      valueCell = sheet.cell(row=row+1,column=column+1).value
      rowArr.append(valueCell)

    if(count % 2 == 0):
      DH = rowArr[6]
      # Calcular Azimut
      grados = rowArr[2]
      minutos = rowArr[3]
      segundos = rowArr[4]
      azimut = round(math.radians(grados+(minutos/60)+(segundos/3600)), 3)
      # Almacenar Azimut en rowArr, para despues guardar en Excel
      rowArr[5] = azimut
      # Calcular Y
      Y = round(math.cos(azimut) * DH, 3)
      X = round(math.sin(azimut) * DH, 3)
      rowArr[7] = Y
      rowArr[8] = X

    rowsExcel.append(rowArr)
    count+=1

  # Calcular Sumatorias DH Y X
  sumDH = 0
  sumY = 0
  sumX = 0
  for row in rowsExcel:
    if (row[0] == None):
      sumDH += row[6]
      sumY += row[7]
      sumX += row[8]

  # Calcular CY - CX - YC - XC
  for row in range(len(rowsExcel)):
    if (rowsExcel[row][0] == None):
      DH = rowsExcel[row][6]
      CY = sumY / sumDH * DH
      CX = sumX / sumDH * DH

      rowsExcel[row][9] = round(CY, 3)
      rowsExcel[row][10] = round(CX, 3)

      YC = rowsExcel[row][7] - CY
      XC = rowsExcel[row][8] - CX

      rowsExcel[row][11] = round(YC, 3)
      rowsExcel[row][12] = round(XC, 3)

  for row in range(len(rowsExcel)):
    if (row == 1):
      rowsExcel[row][13] = round(rowsExcel[row][11] + 2000, 3)
      rowsExcel[row][14] = round(rowsExcel[row][12] + 2000, 3)
    else:
      if (rowsExcel[row][1] != None):
        rowsExcel[row][13] = round(rowsExcel[row][11] + rowsExcel[row - 2][13], 3)
        rowsExcel[row][14] = round(rowsExcel[row][12] + rowsExcel[row - 2][14], 3)

  for row in range(len(rowsExcel)):
    print(rowsExcel[row])

  for row in range(rows-1):
    for column in range(columns):
      if (row >= 3 and rowsExcel[row - 2][1] != None):
        sheet[f'N{row + 1}'] = rowsExcel[row - 2][13]
        sheet[f'O{row + 1}'] = rowsExcel[row - 2][14]

if (hoja == 'Transito'):
  print('Transito')
  sheet = wb['Transito']
  rows = sheet.max_row
  columns = sheet.max_column

  count = 1
  rowsExcel = []

  # FOR's para guardar los datos del Excel en un array
  for row in range(2, rows-1):
    rowArr = []
    for column in range(columns):
      valueCell = sheet.cell(row=row+1,column=column+1).value
      rowArr.append(valueCell)

    if(count % 2 == 0):
      DH = rowArr[6]
      # Calcular Azimut
      grados = rowArr[2]
      minutos = rowArr[3]
      segundos = rowArr[4]
      azimut = round(math.radians(grados+(minutos/60)+(segundos/3600)), 3)
      # Almacenar Azimut en rowArr, para despues guardar en Excel
      rowArr[5] = azimut
      # Calcular Y
      Y = round(math.cos(azimut) * DH, 3)
      X = round(math.sin(azimut) * DH, 3)
      rowArr[7] = Y
      rowArr[9] = X

    rowsExcel.append(rowArr)
    count+=1

  # Calcular Sumatorias DH Y X
  sumDH = 0
  sumY = 0
  sumX = 0
  for row in rowsExcel:
    if (row[0] == None):
      sumDH += row[6]
      sumY += row[7]
      sumX += row[9]

  # Calcular VY VX
  for row in range(len(rowsExcel)):
    if (rowsExcel[row][0] == None):
      rowsExcel[row][8] = abs(rowsExcel[row][7])
      rowsExcel[row][10] = abs(rowsExcel[row][9])

  # Sumatorias VY VX
  sumVY = 0
  sumVX = 0
  for row in rowsExcel:
    if (row[0] == None):
      sumVY += row[8]
      sumVX += row[10]

  # Calcular CY - CX - YC - XC
  for row in range(len(rowsExcel)):
    if (rowsExcel[row][0] == None):
      DH = rowsExcel[row][6]
      CY = (sumY / sumVY) * rowsExcel[row][7]
      CX = (sumX / sumVX) * rowsExcel[row][7]

      rowsExcel[row][11] = round(CY, 3)
      rowsExcel[row][12] = round(CX, 3)

      if (rowsExcel[row][7] < 0):
        YC = (rowsExcel[row][8] - CY) * -1
      else:
        YC = rowsExcel[row][8] - CY

      if (rowsExcel[row][9] < 0):
        XC = (rowsExcel[row][10] - CX) * -1
      else:
        XC = rowsExcel[row][10] - CX

      rowsExcel[row][13] = round(YC, 3)
      rowsExcel[row][14] = round(XC, 3)

  for row in range(len(rowsExcel)):
    if (row == 1):
      rowsExcel[row][15] = round(rowsExcel[row][13] + 2000, 3)
      rowsExcel[row][16] = round(rowsExcel[row][14] + 2000, 3)
    else:
      if (rowsExcel[row][1] != None):
        rowsExcel[row][15] = round(rowsExcel[row][13] + rowsExcel[row - 2][15], 3)
        rowsExcel[row][16] = round(rowsExcel[row][14] + rowsExcel[row - 2][16], 3)

  for row in range(len(rowsExcel)):
    print(rowsExcel[row])

  for row in range(rows-1):
    for column in range(columns):
      if (row >= 3 and rowsExcel[row - 2][1] != None):
        sheet[f'P{row + 1}'] = rowsExcel[row - 2][15]
        sheet[f'Q{row + 1}'] = rowsExcel[row - 2][16]

wb.save('Excel.xlsx')