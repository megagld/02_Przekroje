from pathlib import Path
from Przekroje_funkcje import *
import pandas as pd

# Plik z danymi:
file = '{}{}'.format(Path(__file__).parent,'''/Dane.xlsx''')

# Nazwy arkuszy:
workbook = pd.ExcelFile(file)
sheet_names = workbook.sheet_names

# Rysowanie przekrojów:
x_g = 0
y_g = 0
opis_gora = []

for index, sheet in enumerate(sheet_names):
    # Określenie punktu 0,0:
    print(pd.read_excel(file, sheet_name=sheet).iloc[3,3])
    x_g += pd.read_excel(file, sheet_name=sheet).iloc[3, 3]+index*20
    y_g += pd.read_excel(file, sheet_name=sheet).iloc[4, 3]

    # Rysowanie przekroju ruchowego:
    przekroj_ruchowy = rysowanie_przekroj_ruchowy(x_g, y_g, file, sheet)
    opis_gora.append(przekroj_ruchowy)
    pow_gorna = przekroj_ruchowy[10]

    # Rysowanie konstrukcji:
    if sheet.split('_')[0] == 'B':
        rysowanie_konstrukcja_belkowy(file, sheet, pow_gorna)
    elif sheet.split('_')[0] == 'T':
        rysowanie_konstrukcja_belki_T(file, sheet, pow_gorna)
    elif sheet.split('_')[0] == 'S':
        rysowanie_konstrukcja_skrzynkowy(y_g, file, sheet, pow_gorna)
    elif sheet.split('_')[0] == 'P':
        rysowanie_konstrukcja_plytowy(file, sheet, pow_gorna)
    elif sheet.split('_')[0] == 'Z':
        rysowanie_konstrukcja_zespolony(file, sheet, pow_gorna)
    else:
        print('Zły rodzaj konstrukcji!')
        break

    # Opisywanie:
    if len(sheet.split("_")) == 2 or sheet.split("_")[2] == 'P':
        opis(opis_gora, file, sheet)
        opis_gora = []
