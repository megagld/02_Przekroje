from binhex import LINELEN
from pathlib import Path
import pandas as pd
import Pobieranie_danych

from Przekroj_ruchowy import *
from Opisy import *
from Konstrukcja_plytowo_belkowa import *
from Konstrukcja_zespolona_belki_T import *

# Plik z danymi:
file = '{}{}'.format(Path(__file__).parent,'''/Dane.xlsx''')

# Nazwy arkuszy:
# sheet_names = pd.ExcelFile(file).sheet_names
# ustalone  "na sztywno" jako '01_Zestawienie obiektów'
s_name='01_Zestawienie obiektów'


# Rysowanie przekrojów:
# x_g = 0
# y_g = 0
opis_gora = []

# liczba_obiektów=len(pd.read_excel(file, sheet_name=s_name).columns)
liczba_obiektów=2# -->poskończeniu pisania zamienić na powyższe

# for index, sheet in enumerate(s_name):
for i in range(liczba_obiektów):
    # Pobranie danych
    nr_przekroju=i+5
    Pobieranie_danych.pobierz_dane(file,s_name,nr_przekroju)
    import Pobieranie_danych

    # # Określenie punktu 0,0:
    x_g=Pobieranie_danych.delta_x
    y_g=Pobieranie_danych.delta_y+int(Pobieranie_danych.tom)*-15

    # Rysowanie przekroju ruchowego:
    przekroj_ruchowy = rysowanie_przekroj_ruchowy(x_g, y_g)
    opis_gora.append(przekroj_ruchowy)
    pow_gorna = przekroj_ruchowy[10]

    # Opisywanie:
    opis(opis_gora)
    opis_gora = []


    # # Rysowanie konstrukcji:
    # if sheet.split('_')[0] == 'B':
    if Pobieranie_danych.typ=='płytowo-belkowy':
        rysowanie_konstrukcja_belkowy(pow_gorna)
    elif Pobieranie_danych.typ=='zespolony (belki T)':
        rysowanie_konstrukcja_belki_T(pow_gorna)
    
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


