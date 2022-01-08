import pandas as pd
from pathlib import Path

def pobierz_dane(file,sheet):
    # Pobranie danych:
    global obiekt,delta_y,pasy_lewe,pasy_prawe,awaryjny_lewy,awaryjny_prawy,opaska_lewa,opaska_prawa,chodnik_lewy,chodnik_prawy,bariery_lewa,bariery_prawa,zalamanie_lewe,zalamanie_prawe,konstrukcja
    
    obiekt= 'WD' # do zmiany!!!!
    delta_y = pd.read_excel(file, usecols=[0], sheet_name=sheet).iloc[0, 0]
    pasy_lewe = pd.read_excel(file, usecols=[1, 2, 3, 4], sheet_name=sheet)
    pasy_prawe = pd.read_excel(file, usecols=[5, 6, 7, 8], sheet_name=sheet)
    awaryjny_lewy = pd.read_excel(file, usecols=[9, 10], sheet_name=sheet)
    awaryjny_prawy = pd.read_excel(file, usecols=[11, 12], sheet_name=sheet)
    opaska_lewa = pd.read_excel(file, usecols=[13, 14], sheet_name=sheet)
    opaska_prawa = pd.read_excel(file, usecols=[15, 16], sheet_name=sheet)
    chodnik_lewy = pd.read_excel(file, usecols=[17, 18, 19, 20, 21, 22], sheet_name=sheet)
    chodnik_prawy = pd.read_excel(file, usecols=[23, 24, 25, 26, 27, 28], sheet_name=sheet)
    bariery_lewa = pd.read_excel(file, usecols=[29, 30, 31, 32, 33], sheet_name=sheet)
    bariery_prawa = pd.read_excel(file, usecols=[34, 35, 36, 37, 38], sheet_name=sheet)
    zalamanie_lewe = pd.read_excel(file, usecols=[39, 40], sheet_name=sheet)
    zalamanie_prawe = pd.read_excel(file, usecols=[41, 42], sheet_name=sheet)

    # Pobranie danych:
    konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63], sheet_name=sheet)