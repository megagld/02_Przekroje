from pathlib import Path
import pandas as pd

from Przekroj_ruchowy import *
from Opisy import *
from Konstrukcja_plytowo_belkowa import *

# Plik z danymi:
file = '{}{}'.format(Path(__file__).parent,'''/Dane.xlsx''')

# Nazwy arkuszy:
sheet_names = pd.ExcelFile(file).sheet_names

# Rysowanie przekrojów:
x_g = 0
y_g = 0
opis_gora = []

for index, sheet in enumerate(sheet_names):
    # Pobranie danych
    Pobieranie_danych.pobierz_dane(file,sheet)
    import Pobieranie_danych

    # # Określenie punktu 0,0:
    x_g=Pobieranie_danych.delta_x
    y_g=Pobieranie_danych.delta_y+int(Pobieranie_danych.tom)*-15

    # Rysowanie przekroju ruchowego:
    przekroj_ruchowy = rysowanie_przekroj_ruchowy(x_g, y_g)
    opis_gora.append(przekroj_ruchowy)
    pow_gorna = przekroj_ruchowy[10]

    # Opisywanie:
    opis(opis_gora, sheet)
    opis_gora = []


    # # Rysowanie konstrukcji:
    # if sheet.split('_')[0] == 'B':
    rysowanie_konstrukcja_belkowy(pow_gorna)
    # elif sheet.split('_')[0] == 'T':
    #     rysowanie_konstrukcja_belki_T(file, sheet, pow_gorna)
    # elif sheet.split('_')[0] == 'S':
    #     rysowanie_konstrukcja_skrzynkowy(y_g, file, sheet, pow_gorna)
    # elif sheet.split('_')[0] == 'P':
    #     rysowanie_konstrukcja_plytowy(file, sheet, pow_gorna)
    # elif sheet.split('_')[0] == 'Z':
    #     rysowanie_konstrukcja_zespolony(file, sheet, pow_gorna)
    # else:
    #     print('Zły rodzaj konstrukcji!')
    #     break


