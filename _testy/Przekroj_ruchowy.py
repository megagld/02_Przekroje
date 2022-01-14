from pyautocad import aDouble, Autocad
import pandas as pd
from math import isnan, atan, fabs, radians, sin, cos
from itertools import chain
import win32com.client
import Pobieranie_danych
from Funkcje_podstawowe import *


acad = Autocad()

version = acad.doc.GetVariable("ACADVER")
color = acad.app.GetInterfaceObject(f'AutoCAD.ACCmColor.{version[0:2]}')
acad_32 = win32com.client.Dispatch("AutoCAD.Application")


@speed_test
def rysowanie_przekroj_ruchowy(x_g, y_g):

    # Pobranie danych:
    # delta_y = pd.read_excel(file, usecols=[0], sheet_name=sheet).iloc[0, 0]
    # pasy_lewe = pd.read_excel(file, usecols=[1, 2, 3, 4], sheet_name=sheet)
    # pasy_prawe = pd.read_excel(file, usecols=[5, 6, 7, 8], sheet_name=sheet)            
    # awaryjny_lewy = pd.read_excel(file, usecols=[9, 10], sheet_name=sheet)
    # awaryjny_prawy = pd.read_excel(file, usecols=[11, 12], sheet_name=sheet)
    # opaska_lewa = pd.read_excel(file, usecols=[13, 14], sheet_name=sheet)
    # opaska_prawa = pd.read_excel(file, usecols=[15, 16], sheet_name=sheet)
    # chodnik_lewy = pd.read_excel(file, usecols=[17, 18, 19, 20, 21, 22], sheet_name=sheet)
    # chodnik_prawy = pd.read_excel(file, usecols=[23, 24, 25, 26, 27, 28], sheet_name=sheet)
    # bariery_lewa = pd.read_excel(file, usecols=[29, 30, 31, 32, 33], sheet_name=sheet)
    # bariery_prawa = pd.read_excel(file, usecols=[34, 35, 36, 37, 38], sheet_name=sheet)
    # zalamanie_lewe = pd.read_excel(file, usecols=[39, 40], sheet_name=sheet)
    # zalamanie_prawe = pd.read_excel(file, usecols=[41, 42], sheet_name=sheet)
        
    global obiekt
    obiekt = Pobieranie_danych.obiekt
    delta_y = Pobieranie_danych.delta_y
    pasy_lewe = Pobieranie_danych.pasy_lewe
    pasy_prawe = Pobieranie_danych.pasy_prawe
    awaryjny_lewy = Pobieranie_danych.awaryjny_lewy
    awaryjny_prawy = Pobieranie_danych.awaryjny_prawy
    opaska_lewa = Pobieranie_danych.opaska_lewa
    opaska_prawa = Pobieranie_danych.opaska_prawa
    chodnik_lewy = Pobieranie_danych.chodnik_lewy
    chodnik_prawy = Pobieranie_danych.chodnik_prawy
    bariery_lewa = Pobieranie_danych.bariery_lewa
    bariery_prawa = Pobieranie_danych.bariery_prawa
    zalamanie_lewe = Pobieranie_danych.zalamanie_lewe
    zalamanie_prawe = Pobieranie_danych.zalamanie_prawe
    konstrukcja = Pobieranie_danych.konstrukcja

    # ==================================================================================================================
    # JEZDNIA
    # ==================================================================================================================

    # Punkt 0,0:
    jezdnia = [[x_g, y_g]]

    # Punkty:
    # for index, row in pasy_lewe.iterrows():
    #     if index == 0:
    #         x = x_g
    #         y = y_g
    #     if isnan(pasy_lewe.iloc[index, 0]) or pasy_lewe.iloc[index, 0] == 0:
    #         break
    #     x -= row['PL - szer']
    #     y += row['PL - szer'] * row['PL - spadek'] / 100
    #     jezdnia.append([round(x, 8), round(y, 8)])

    x = x_g
    y = y_g

    for width in pasy_lewe['PL - szer'].split('+'):            
        x -= float(width)
        y += float(width) * pasy_lewe['PL - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    for index, row in awaryjny_lewy.iterrows():
        if isnan(awaryjny_lewy.iloc[index, 0]) or awaryjny_lewy.iloc[index, 0] == 0:
            break
        x -= row['PAL - szer']
        y += row['PAL - szer'] * row['PAL - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    for index, row in opaska_lewa.iterrows():
        if isnan(opaska_lewa.iloc[index, 0]) or opaska_lewa.iloc[index, 0] == 0:
            break
        x -= row['OL - szer']
        y += row['OL - szer'] * row['OL - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    for index, row in pasy_prawe.iterrows():
        if index == 0:
            x = x_g
            y = y_g
        if isnan(pasy_prawe.iloc[index, 0]) or pasy_prawe.iloc[index, 0] == 0:
            break
        x += row['PP - szer']
        y += row['PP - szer'] * row['PP - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    for index, row in awaryjny_prawy.iterrows():
        if isnan(awaryjny_prawy.iloc[index, 0]) or awaryjny_prawy.iloc[index, 0] == 0:
            break
        x += row['PAP - szer']
        y += row['PAP - szer'] * row['PAP - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    for index, row in opaska_prawa.iterrows():
        if isnan(opaska_prawa.iloc[index, 0]) or opaska_prawa.iloc[index, 0] == 0:
            break
        x += row['OP - szer']
        y += row['OP - szer'] * row['OP - spadek'] / 100
        jezdnia.append([round(x, 8), round(y, 8)])

    # Rysowanie pierwszej linii:
    jezdnia = sorted(jezdnia)
    jezdnia_lista = list(chain.from_iterable(jezdnia))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(jezdnia_lista))
    LWPline.Layer = 'AII_M_nawierzchnia'

    # Rysowanie drugiej linii:
    jezdnia_2 = [[i[0], round(i[1] - 0.04, 8)] for i in jezdnia]
    jezdnia_2_del = []
    for i in range(len(jezdnia_2) - 2):
        xi = jezdnia_2[i + 1][0]
        yi = jezdnia_2[i + 1][1]
        x1 = jezdnia_2[i][0]
        y1 = jezdnia_2[i][1]
        x2 = jezdnia_2[i + 2][0]
        y2 = jezdnia_2[i + 2][1]
        if round((y2 - yi) / (x2 - xi), 6) == round((yi - y1) / (xi - x1), 6):
            jezdnia_2_del.append([xi, yi])
    for i in jezdnia_2_del:
        jezdnia_2.remove(i)

    jezdnia_2_lista = list(chain.from_iterable(jezdnia_2))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(jezdnia_2_lista))
    LWPline.Layer = 'AII_M_nawierzchnia'
    color.ColorIndex = 8
    LWPline.TrueColor = color

    # ==================================================================================================================
    # KRAWĘŻNIKI
    # ==================================================================================================================

    # Krawężnik lewy:

    # Punkty:
    x_kraw_l = jezdnia[0][0]
    y_kraw_l = jezdnia[0][1]
    kraw_lewy = [x_kraw_l, y_kraw_l - 0.04, x_kraw_l - 0.2, y_kraw_l - 0.04, x_kraw_l - 0.2, y_kraw_l + 0.14,
                 x_kraw_l - 0.04, y_kraw_l + 0.14, x_kraw_l, y_kraw_l + 0.04]

    # Rysowanie:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(kraw_lewy))
    LWPline.Closed = True
    LWPline.Layer = 'AII_M_krawężnik'

    # ------------------------------------------------------------------------------------------------------------------
    # Krawężnik prawy:

    # Punkty:
    x_kraw_p = jezdnia[-1][0]
    y_kraw_p = jezdnia[-1][1]
    kraw_prawy = [x_kraw_p, y_kraw_p - 0.04, x_kraw_p + 0.2, y_kraw_p - 0.04, x_kraw_p + 0.2, y_kraw_p + 0.14,
                  x_kraw_p + 0.04, y_kraw_p + 0.14, x_kraw_p, y_kraw_p + 0.04]

    # Rysowanie:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(kraw_prawy))
    LWPline.Closed = True
    LWPline.Layer = 'AII_M_krawężnik'

    print('KRAWĘŻNIKI')
    
    # ==================================================================================================================
    # CHODNIKI
    # ==================================================================================================================

    # Chodnik lewy:

    # Pobranie danych:
    ch_lewy_szer_ch = chodnik_lewy.iloc[0, 0]
    ch_lewy_szer_sr = chodnik_lewy.iloc[0, 1]
    ch_lewy_szer_cpr = chodnik_lewy.iloc[0, 2]
    ch_lewy_szer = ch_lewy_szer_ch + ch_lewy_szer_sr + ch_lewy_szer_cpr
    ch_lewy_sp = chodnik_lewy.iloc[0, 3]
    ch_lewy_deska = chodnik_lewy.iloc[0, 4]
    ch_lewy_lat = chodnik_lewy.iloc[0, 5]
    bar_lewa = bariery_lewa.iloc[0, 0]
    bar_lewa_rodz = bariery_lewa.iloc[0, 1]
    bar_lewa_op = bariery_lewa.iloc[0, 2]
    bal_lewa_rodz = bariery_lewa.iloc[0, 3]
    bal_lewa_wys = bariery_lewa.iloc[0, 4]

    # do poprawy !!!!- powinno pobierać dane jako liczbę, a nie jako str
    bal_lewa_wys=float(bal_lewa_wys)    

    # Punkt początkowy:
    x_ch_lewy_1 = x_kraw_l - 0.2
    y_ch_lewy_1 = y_kraw_l + 0.14

    # Określenie szerokości:
    if bar_lewa_rodz == 'bariera linowa':
        bar_lewa_szer = 0.15
    else:
        bar_lewa_szer = 0.40

    if bal_lewa_rodz[0] == 'b':
        bal_lewa_szer = 0.21
    else:
        bal_lewa_szer = 0.41

    if bar_lewa == 'T':
        delta_x_ch_lewy = bar_lewa_op + bar_lewa_szer + ch_lewy_szer + bal_lewa_szer - 0.20
    else:
        if ch_lewy_szer == 0:
            delta_x_ch_lewy = bar_lewa_op + 0.56 - 0.20
        else:
            delta_x_ch_lewy = ch_lewy_szer + 0.56 - 0.20

    
    # Punkt końcowy:
    x_ch_lewy_2 = x_ch_lewy_1 - delta_x_ch_lewy
    y_ch_lewy_2 = round(y_ch_lewy_1 + delta_x_ch_lewy * ch_lewy_sp / 100, 8)

    # Rysowanie:
    chod_lewy = [x_ch_lewy_2, y_ch_lewy_2, x_ch_lewy_1, y_ch_lewy_1]
    LWPline = acad.model.AddLightWeightPolyline(aDouble(chod_lewy))
    LWPline.Layer = 'AII_M_nawierzchnia'

    # Wstawienie bloku deski:
    Deska_lewa = acad.model.InsertBlock(aDouble(x_ch_lewy_2, y_ch_lewy_2, 0), f'Deska_{ch_lewy_deska}', -1, 1, 1, 0)
    Deska_lewa.Layer = 'AII_M_wyposażenie'

    # Wstawianie bloku bariery:
    if bar_lewa == 'T':
        if bar_lewa_rodz == 'bariera linowa':
            x_bar_lewa = x_ch_lewy_1 + 0.20 - bar_lewa_op - 0.075
            y_bar_lewa = round(y_ch_lewy_1 + (bar_lewa_op + 0.1625 - 0.20) * ch_lewy_sp / 100 + 0.01, 8)
            name = 'Bariera_linowa'
            x_bar_lewa_2 = round(x_bar_lewa - 0.0875, 6)
            y_bar_lewa_2 = y_bar_lewa
            x_bar_lewa_1 = round((x_bar_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bar_lewa_1 = round(y_bar_lewa_2 + x_bar_lewa_1 - x_bar_lewa_2, 12)
            x_bar_lewa_3 = round(x_bar_lewa + 0.0875, 6)
            y_bar_lewa_3 = y_bar_lewa
            x_bar_lewa_4 = round(
                (x_bar_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.175 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bar_lewa_4 = round(y_bar_lewa_3 - x_bar_lewa_4 + x_bar_lewa_3, 12)
        else:
            x_bar_lewa = x_ch_lewy_1 + 0.20 - bar_lewa_op - 0.25
            y_bar_lewa = round(y_ch_lewy_1 + (bar_lewa_op + 0.37 - 0.20) * ch_lewy_sp / 100 + 0.01, 8)
            name = 'Bariera'
            x_bar_lewa_2 = round(x_bar_lewa - 0.12, 6)
            y_bar_lewa_2 = y_bar_lewa
            x_bar_lewa_1 = round((x_bar_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bar_lewa_1 = round(y_bar_lewa_2 + x_bar_lewa_1 - x_bar_lewa_2, 12)
            x_bar_lewa_3 = round(x_bar_lewa + 0.12, 6)
            y_bar_lewa_3 = y_bar_lewa
            x_bar_lewa_4 = round(
                (x_bar_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.24 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bar_lewa_4 = round(y_bar_lewa_3 - x_bar_lewa_4 + x_bar_lewa_3, 12)
    else:
        if ch_lewy_szer == 0:
            x_bar_lewa = x_ch_lewy_1 + 0.20 - bar_lewa_op - 0.31
            y_bar_lewa = round(y_ch_lewy_1 + (bar_lewa_op + 0.52 - 0.20) * ch_lewy_sp / 100 + 0.01, 8)
            name = f'Barieroporęcz_{bar_lewa_rodz[-3:]}'
            x_bar_lewa_2 = round(x_bar_lewa - 0.21, 6)
            y_bar_lewa_2 = y_bar_lewa
            x_bar_lewa_1 = round((x_bar_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bar_lewa_1 = round(y_bar_lewa_2 + x_bar_lewa_1 - x_bar_lewa_2, 12)
            x_bar_lewa_3 = round(x_bar_lewa + 0.21, 6)
            y_bar_lewa_3 = y_bar_lewa
            x_bar_lewa_4 = round(
                (x_bar_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.42 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bar_lewa_4 = round(y_bar_lewa_3 - x_bar_lewa_4 + x_bar_lewa_3, 12)
        else:
            x_bar_lewa = x_ch_lewy_1 + 0.20 - ch_lewy_szer - 0.31
            y_bar_lewa = round(y_ch_lewy_1 + (ch_lewy_szer + 0.52 - 0.20) * ch_lewy_sp / 100 + 0.01, 8)
            name = f'Barieroporęcz_{bar_lewa_rodz[-3:]}'
            x_bar_lewa_2 = round(x_bar_lewa - 0.21, 6)
            y_bar_lewa_2 = y_bar_lewa
            x_bar_lewa_1 = round((x_bar_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bar_lewa_1 = round(y_bar_lewa_2 + x_bar_lewa_1 - x_bar_lewa_2, 12)
            x_bar_lewa_3 = round(x_bar_lewa + 0.21, 6)
            y_bar_lewa_3 = y_bar_lewa
            x_bar_lewa_4 = round(
                (x_bar_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.42 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bar_lewa_4 = round(y_bar_lewa_3 - x_bar_lewa_4 + x_bar_lewa_3, 12)

    # Wstawianie bloku:
    Bar_lewa = acad.model.InsertBlock(aDouble(x_bar_lewa, y_bar_lewa, 0), name, -1, 1, 1, 0)
    Bar_lewa.Layer = 'AII_M_bariery'

    # Rysowanie podlewki:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bar_lewa_1, y_bar_lewa_1, x_bar_lewa_2, y_bar_lewa_2))
    LWPline.Layer = 'AII_M_bariery'
    color.ColorIndex = 8
    LWPline.TrueColor = color
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bar_lewa_3, y_bar_lewa_3, x_bar_lewa_4, y_bar_lewa_4))
    LWPline.Layer = 'AII_M_bariery'
    color.ColorIndex = 8
    LWPline.TrueColor = color

    # Wstawianie balustrady/ekranu:
    if bar_lewa == 'T':
        if bal_lewa_rodz[0] == 'b':
            x_bal_lewa = x_ch_lewy_2 + 0.17
            y_bal_lewa = round(y_ch_lewy_2 - 0.11 * ch_lewy_sp / 100 + 0.01, 8)
            name = f'Balustrada_{bal_lewa_wys}'
            x_bal_lewa_2 = round(x_bal_lewa - 0.06, 6)
            y_bal_lewa_2 = y_bal_lewa
            x_bal_lewa_1 = round((x_bal_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bal_lewa_1 = round(y_bal_lewa_2 + x_bal_lewa_1 - x_bal_lewa_2, 12)
            x_bal_lewa_3 = round(x_bal_lewa + 0.06, 6)
            y_bal_lewa_3 = y_bal_lewa
            x_bal_lewa_4 = round(
                (x_bal_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.12 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bal_lewa_4 = round(y_bal_lewa_3 - x_bal_lewa_4 + x_bal_lewa_3, 12)

        else:
            x_bal_lewa = x_ch_lewy_2 + 0.285
            y_bal_lewa = round(y_ch_lewy_2 - 0.06 * ch_lewy_sp / 100 + 0.01, 8)
            name = f'Ekran_{bal_lewa_wys}'
            x_bal_lewa_2 = round(x_bal_lewa - 0.285, 6)
            y_bal_lewa_2 = y_bal_lewa
            x_bal_lewa_1 = round((x_bal_lewa_2 * (1 + ch_lewy_sp / 100) - 0.01) / (1 + ch_lewy_sp / 100), 12)
            y_bal_lewa_1 = round(y_bal_lewa_2 + x_bal_lewa_1 - x_bal_lewa_2, 12)
            x_bal_lewa_3 = round(x_bal_lewa + 0.175, 6)
            y_bal_lewa_3 = y_bal_lewa
            x_bal_lewa_4 = round(
                (x_bal_lewa_3 * (1 - ch_lewy_sp / 100) + 0.01 + 0.35 * ch_lewy_sp / 100) / (1 - ch_lewy_sp / 100), 12)
            y_bal_lewa_4 = round(y_bal_lewa_3 - x_bal_lewa_4 + x_bal_lewa_3, 12)

        # Wstawianie bloku:
        Bal_lewa = acad.model.InsertBlock(aDouble(x_bal_lewa, y_bal_lewa, 0), name, -1, 1, 1, 0)
        Bal_lewa.Layer = 'AII_M_balustrady'

        # Rysowanie podlewki:
        LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bal_lewa_1, y_bal_lewa_1, x_bal_lewa_2, y_bal_lewa_2))
        LWPline.Layer = 'AII_M_balustrady'
        color.ColorIndex = 8
        LWPline.TrueColor = color
        LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bal_lewa_3, y_bal_lewa_3, x_bal_lewa_4, y_bal_lewa_4))
        LWPline.Layer = 'AII_M_balustrady'
        color.ColorIndex = 8
        LWPline.TrueColor = color

    # Wstawianie bloku latarni:
    if ch_lewy_lat == 'T':
        if ch_lewy_sp % 1 == 0:
            name = f'Latarnia_{ch_lewy_deska}_{int(ch_lewy_sp)}'
        else:
            name = f'Latarnia_{ch_lewy_deska}_{ch_lewy_sp}'
        Lat_lewa = acad.model.InsertBlock(aDouble(x_ch_lewy_2, y_ch_lewy_2, 0), name, -1, 1, 1, 0)
        Lat_lewa.Layer = 'AII_M_wyposażenie'
    # ------------------------------------------------------------------------------------------------------------------
    # Chodnik prawy:

    # Pobranie danych:
    ch_prawy_szer_ch = chodnik_prawy.iloc[0, 0]
    ch_prawy_szer_sr = chodnik_prawy.iloc[0, 1]
    ch_prawy_szer_cpr = chodnik_prawy.iloc[0, 2]
    ch_prawy_szer = ch_prawy_szer_ch + ch_prawy_szer_sr + ch_prawy_szer_cpr
    ch_prawy_sp = chodnik_prawy.iloc[0, 3]
    ch_prawy_deska = chodnik_prawy.iloc[0, 4]
    ch_prawy_lat = chodnik_prawy.iloc[0, 5]
    bar_prawa = bariery_prawa.iloc[0, 0]
    bar_prawa_rodz = bariery_prawa.iloc[0, 1]
    bar_prawa_op = bariery_prawa.iloc[0, 2]
    bal_prawa_rodz = bariery_prawa.iloc[0, 3]
    bal_prawa_wys = bariery_prawa.iloc[0, 4]

    # do poprawy !!!!- powinno pobierać dane jako liczbę, a nie jako str
    bal_prawa_wys=float(bal_prawa_wys)   

    # Punkt początkowy:
    x_ch_prawy_1 = x_kraw_p + 0.2
    y_ch_prawy_1 = y_kraw_p + 0.14

    # Określenie szerokości:
    if bar_prawa_rodz == 'bariera linowa':
        bar_prawa_szer = 0.15
    else:
        bar_prawa_szer = 0.40

    if bal_prawa_rodz[0] == 'b':
        bal_prawa_szer = 0.21
    else:
        bal_prawa_szer = 0.41

    if bar_prawa == 'T':
        delta_x_ch_prawy = bar_prawa_op + bar_prawa_szer + ch_prawy_szer + bal_prawa_szer - 0.20
    else:
        if ch_prawy_szer == 0:
            delta_x_ch_prawy = bar_prawa_op + 0.56 - 0.20
        else:
            delta_x_ch_prawy = ch_prawy_szer + 0.56 - 0.20

    # Punkt końcowy:
    x_ch_prawy_2 = x_ch_prawy_1 + delta_x_ch_prawy
    y_ch_prawy_2 = round(y_ch_prawy_1 + delta_x_ch_prawy * ch_prawy_sp / 100, 8)

    # Rysowanie:
    chod_prawy = [x_ch_prawy_1, y_ch_prawy_1, x_ch_prawy_2, y_ch_prawy_2]
    LWPline = acad.model.AddLightWeightPolyline(aDouble(chod_prawy))
    LWPline.Layer = 'AII_M_nawierzchnia'

    # Wstawienie bloku deski:
    Deska_prawa = acad.model.InsertBlock(aDouble(x_ch_prawy_2, y_ch_prawy_2, 0), f'Deska_{ch_prawy_deska}', 1, 1, 1, 0)
    Deska_prawa.Layer = 'AII_M_wyposażenie'

    # Wstawianie bloku bariery:
    if bar_prawa == 'T':
        if bar_prawa_rodz == 'bariera linowa':
            x_bar_prawa = x_ch_prawy_1 - 0.20 + bar_prawa_op + 0.075
            y_bar_prawa = round(y_ch_prawy_1 + (bar_prawa_op + 0.1625 - 0.20) * ch_prawy_sp / 100 + 0.01, 8)
            name = 'Bariera_linowa'
            x_bar_prawa_3 = round(x_bar_prawa + 0.0875, 6)
            y_bar_prawa_3 = y_bar_prawa
            x_bar_prawa_4 = round((x_bar_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bar_prawa_4 = round(y_bar_prawa_3 - x_bar_prawa_4 + x_bar_prawa_3, 12)
            x_bar_prawa_2 = round(x_bar_prawa - 0.0875, 6)
            y_bar_prawa_2 = y_bar_prawa
            x_bar_prawa_1 = round(
                (x_bar_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.175 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bar_prawa_1 = round(y_bar_prawa_2 + x_bar_prawa_1 - x_bar_prawa_2, 12)
        else:
            x_bar_prawa = x_ch_prawy_1 - 0.20 + bar_prawa_op + 0.25
            y_bar_prawa = round(y_ch_prawy_1 + (bar_prawa_op + 0.37 - 0.20) * ch_prawy_sp / 100 + 0.01, 8)
            name = 'Bariera'
            x_bar_prawa_3 = round(x_bar_prawa + 0.12, 6)
            y_bar_prawa_3 = y_bar_prawa
            x_bar_prawa_4 = round((x_bar_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bar_prawa_4 = round(y_bar_prawa_3 - x_bar_prawa_4 + x_bar_prawa_3, 12)
            x_bar_prawa_2 = round(x_bar_prawa - 0.12, 6)
            y_bar_prawa_2 = y_bar_prawa
            x_bar_prawa_1 = round(
                (x_bar_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.24 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bar_prawa_1 = round(y_bar_prawa_2 + x_bar_prawa_1 - x_bar_prawa_2, 12)
    else:
        if ch_prawy_szer == 0:
            x_bar_prawa = x_ch_prawy_1 - 0.20 + bar_prawa_op + 0.31
            y_bar_prawa = round(y_ch_prawy_1 + (bar_prawa_op + 0.52 - 0.20) * ch_prawy_sp / 100 + 0.01, 8)
            name = f'Barieroporęcz_{bar_prawa_rodz[-3:]}'
            x_bar_prawa_3 = round(x_bar_prawa + 0.21, 6)
            y_bar_prawa_3 = y_bar_prawa
            x_bar_prawa_4 = round((x_bar_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bar_prawa_4 = round(y_bar_prawa_3 - x_bar_prawa_4 + x_bar_prawa_3, 12)
            x_bar_prawa_2 = round(x_bar_prawa - 0.21, 6)
            y_bar_prawa_2 = y_bar_prawa
            x_bar_prawa_1 = round(
                (x_bar_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.42 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bar_prawa_1 = round(y_bar_prawa_2 + x_bar_prawa_1 - x_bar_prawa_2, 12)
        else:
            x_bar_prawa = x_ch_prawy_1 - 0.20 + ch_prawy_szer + 0.31
            y_bar_prawa = round(y_ch_prawy_1 + (ch_prawy_szer + 0.52 - 0.20) * ch_prawy_sp / 100 + 0.01, 8)
            name = f'Barieroporęcz_{bar_prawa_rodz[-3:]}'
            x_bar_prawa_3 = round(x_bar_prawa + 0.21, 6)
            y_bar_prawa_3 = y_bar_prawa
            x_bar_prawa_4 = round((x_bar_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bar_prawa_4 = round(y_bar_prawa_3 - x_bar_prawa_4 + x_bar_prawa_3, 12)
            x_bar_prawa_2 = round(x_bar_prawa - 0.21, 6)
            y_bar_prawa_2 = y_bar_prawa
            x_bar_prawa_1 = round(
                (x_bar_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.42 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bar_prawa_1 = round(y_bar_prawa_2 + x_bar_prawa_1 - x_bar_prawa_2, 12)

    # Wstawianie bloku:
    Bar_prawa = acad.model.InsertBlock(aDouble(x_bar_prawa, y_bar_prawa, 0), name, 1, 1, 1, 0)
    Bar_prawa.Layer = 'AII_M_bariery'

    # Rysowanie podlewki:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bar_prawa_1, y_bar_prawa_1, x_bar_prawa_2, y_bar_prawa_2))
    LWPline.Layer = 'AII_M_bariery'
    color.ColorIndex = 8
    LWPline.TrueColor = color
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bar_prawa_3, y_bar_prawa_3, x_bar_prawa_4, y_bar_prawa_4))
    LWPline.Layer = 'AII_M_bariery'
    color.ColorIndex = 8
    LWPline.TrueColor = color

    # Wstawianie balustrady/ekranu:
    if bar_prawa == 'T':
        if bal_prawa_rodz[0] == 'b':
            x_bal_prawa = x_ch_prawy_2 - 0.17
            y_bal_prawa = round(y_ch_prawy_2 - 0.11 * ch_prawy_sp / 100 + 0.01, 8)
            name = f'Balustrada_{bal_prawa_wys}'
            x_bal_prawa_3 = round(x_bal_prawa + 0.06, 6)
            y_bal_prawa_3 = y_bal_prawa
            x_bal_prawa_4 = round((x_bal_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bal_prawa_4 = round(y_bal_prawa_3 - x_bal_prawa_4 + x_bal_prawa_3, 12)
            x_bal_prawa_2 = round(x_bal_prawa - 0.06, 6)
            y_bal_prawa_2 = y_bal_prawa
            x_bal_prawa_1 = round(
                (x_bal_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.12 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bal_prawa_1 = round(y_bal_prawa_2 + x_bal_prawa_1 - x_bal_prawa_2, 12)
        else:
            x_bal_prawa = x_ch_prawy_2 - 0.285
            y_bal_prawa = round(y_ch_prawy_2 - 0.06 * ch_prawy_sp / 100 + 0.01, 8)
            name = f'Ekran_{bal_prawa_wys}'
            x_bal_prawa_3 = round(x_bal_prawa + 0.235, 6)
            y_bal_prawa_3 = y_bal_prawa
            x_bal_prawa_4 = round((x_bal_prawa_3 * (1 + ch_prawy_sp / 100) + 0.01) / (1 + ch_prawy_sp / 100), 12)
            y_bal_prawa_4 = round(y_bal_prawa_3 - x_bal_prawa_4 + x_bal_prawa_3, 12)
            x_bal_prawa_2 = round(x_bal_prawa - 0.175, 6)
            y_bal_prawa_2 = y_bal_prawa
            x_bal_prawa_1 = round(
                (x_bal_prawa_2 * (1 - ch_prawy_sp / 100) - 0.01 - 0.35 * ch_prawy_sp / 100) / (1 - ch_prawy_sp / 100),
                12)
            y_bal_prawa_1 = round(y_bal_prawa_2 + x_bal_prawa_1 - x_bal_prawa_2, 12)

        # Wstawianie bloku:
        Bal_prawa = acad.model.InsertBlock(aDouble(x_bal_prawa, y_bal_prawa, 0), name, 1, 1, 1, 0)
        Bal_prawa.Layer = 'AII_M_balustrady'

        # Rysowanie podlewki:
        LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bal_prawa_1, y_bal_prawa_1, x_bal_prawa_2, y_bal_prawa_2))
        LWPline.Layer = 'AII_M_balustrady'
        color.ColorIndex = 8
        LWPline.TrueColor = color
        LWPline = acad.model.AddLightWeightPolyline(aDouble(x_bal_prawa_3, y_bal_prawa_3, x_bal_prawa_4, y_bal_prawa_4))
        LWPline.Layer = 'AII_M_balustrady'
        color.ColorIndex = 8
        LWPline.TrueColor = color

    # Wstawianie bloku latarni:
    if ch_prawy_lat == 'T':
        if ch_prawy_sp % 1 == 0:
            name = f'Latarnia_{ch_prawy_deska}_{int(ch_prawy_sp)}'
        else:
            name = f'Latarnia_{ch_prawy_deska}_{ch_prawy_sp}'
        Lat_prawa = acad.model.InsertBlock(aDouble(x_ch_prawy_2, y_ch_prawy_2, 0), name, 1, 1, 1, 0)
        Lat_prawa.Layer = 'AII_M_wyposażenie'

    print('CHODNIKI')

    # ==================================================================================================================
    # IZOLACJA
    # ==================================================================================================================

    jezdnia_3 = [(i[0], round(i[1] - 0.04, 8)) for i in jezdnia_2]

    # Wspornik lewy:

    # Krawędź jezdni:
    i_lewy_1 = jezdnia_3[0]
    i_lewy_2 = jezdnia_3[1]

    # Spadek pod chodnikiem:
    if zalamanie_lewe.iloc[0, 0] == 'T':
        wsp_lewy_spadek = round(ch_lewy_sp / 100, 8)
    else:
        wsp_lewy_spadek = round((i_lewy_1[1] - i_lewy_2[1]) / (i_lewy_2[0] - i_lewy_1[0]), 8)

    # Przesunięcie punktu załamania:
    delta_x_i_lewy = zalamanie_lewe.iloc[0, 1]

    # Punkt załamania:
    x_i_lewy = i_lewy_1[0] + delta_x_i_lewy
    y_i_lewy = round(i_lewy_1[1] + (x_i_lewy - i_lewy_1[0]) * (i_lewy_2[1] - i_lewy_1[1]) / (i_lewy_2[0] - i_lewy_1[0]),
                     8)

    # Punkt końcowy:
    x_i_lewy_2 = x_ch_lewy_2
    y_i_lewy_2 = round(y_i_lewy + (x_i_lewy - x_i_lewy_2) * wsp_lewy_spadek, 8)

    # ------------------------------------------------------------------------------------------------------------------
    # Wspornik prawy:

    # Krawędź jezdni:
    i_prawy_1 = jezdnia_3[-1]
    i_prawy_2 = jezdnia_3[-2]

    # Spadek pod chodnikiem:
    if zalamanie_prawe.iloc[0, 0] == 'T':
        wsp_prawy_spadek = round(ch_prawy_sp / 100, 8)
    else:
        wsp_prawy_spadek = round((i_prawy_1[1] - i_prawy_2[1]) / (i_prawy_1[0] - i_prawy_2[0]), 8)

    # Przesunięcie punktu załamania:
    delta_x_i_prawy = zalamanie_prawe.iloc[0, 1]

    # Punkt załamania:
    x_i_prawy = i_prawy_1[0] - delta_x_i_prawy
    y_i_prawy = round(
        i_prawy_1[1] + (i_prawy_1[0] - x_i_prawy) * (i_prawy_2[1] - i_prawy_1[1]) / (i_prawy_1[0] - i_prawy_2[0]), 8)

    # Punkt końcowy:
    x_i_prawy_2 = x_ch_prawy_2
    y_i_prawy_2 = round(y_i_prawy + (x_i_prawy_2 - x_i_prawy) * wsp_prawy_spadek, 8)

    # ------------------------------------------------------------------------------------------------------------------
    # Rysowanie górnej linii:
    jezdnia_3 = jezdnia_3[1:-1]
    if zalamanie_lewe.iloc[0, 0] == 'T':
        jezdnia_3.append((x_i_lewy, y_i_lewy))
    jezdnia_3.append((x_i_lewy_2, y_i_lewy_2))
    if zalamanie_prawe.iloc[0, 0] == 'T':
        jezdnia_3.append((x_i_prawy, y_i_prawy))
    jezdnia_3.append((x_i_prawy_2, y_i_prawy_2))
    jezdnia_3 = sorted(jezdnia_3)
    jezdnia_3_lista = list(chain.from_iterable(jezdnia_3))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(jezdnia_3_lista))
    LWPline.Layer = 'AII_M_izolacja'

    # ------------------------------------------------------------------------------------------------------------------
    # Wygięcie izolacji w dół:

    # Wysokość wspornika i współrzędne:
    delta_y_i = konstrukcja.iloc[0, 0]
    i_k_lewy_1 = jezdnia_3[0]
    i_k_lewy_2 = jezdnia_3[1]
    i_k_prawy_1 = jezdnia_3[-1]
    i_k_prawy_2 = jezdnia_3[-2]

    # Punkty:
    x_i_k_lewy = round(i_k_lewy_1[0] + 0.005, 8)
    x_i_k_prawy = round(i_k_prawy_1[0] - 0.005, 8)
    y_i_k_lewy_1 = round(y_i_lewy_2 - 0.005 + (x_i_k_lewy - i_k_lewy_1[0]) * (i_k_lewy_2[1] - i_k_lewy_1[1]) / (
            i_k_lewy_2[0] - i_k_lewy_1[0]), 8)
    y_i_k_lewy_2 = round(
        y_i_lewy_2 - 0.01 - delta_y_i + 2 * (x_i_k_lewy - i_k_lewy_1[0]) * (i_k_lewy_2[1] - i_k_lewy_1[1]) / (
                i_k_lewy_2[0] - i_k_lewy_1[0]), 8)
    y_i_k_prawy_1 = round(y_i_prawy_2 - 0.005 + (x_i_k_prawy - i_k_prawy_1[0]) * (i_k_prawy_2[1] - i_k_prawy_1[1]) / (
            i_k_prawy_2[0] - i_k_prawy_1[0]), 8)
    y_i_k_prawy_2 = round(
        y_i_prawy_2 - 0.01 - delta_y_i + 2 * (x_i_k_prawy - i_k_prawy_1[0]) * (i_k_prawy_2[1] - i_k_prawy_1[1]) / (
                i_k_prawy_2[0] - i_k_prawy_1[0]), 8)

    # Rysowanie linii izolacji:
    jezdnia_4 = [(i[0], round(i[1] - 0.005, 8)) for i in jezdnia_3[1:-1]]
    jezdnia_4.insert(0, [x_i_k_lewy, y_i_k_lewy_2])
    jezdnia_4.insert(1, [x_i_k_lewy, y_i_k_lewy_1])
    jezdnia_4.append([x_i_k_prawy, y_i_k_prawy_1])
    jezdnia_4.append([x_i_k_prawy, y_i_k_prawy_2])
    jezdnia_4_lista = list(chain.from_iterable(jezdnia_4))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(jezdnia_4_lista))
    LWPline.Layer = 'AII_M_izolacja'
    LWPline.Linetype = 'Hidden'
    LWPline.LinetypeScale = 0.7
    LWPline.LinetypeGeneration = True
    LWPline.ConstantWidth = 0.01

    print('IZOLACJA')

    # ==================================================================================================================
    # PODLEWKI POD KRAWĘŻNIKI
    # ==================================================================================================================

    # Krawężnik lewy:
    x_pk_lewy_2 = kraw_lewy[2]
    y_pk_lewy_2 = kraw_lewy[3]
    x_pk_lewy_3 = kraw_lewy[0]
    y_pk_lewy_3 = kraw_lewy[1]

    # Punkty podlewki:
    x_pk_lewy_1 = round((y_i_lewy + x_i_lewy * wsp_lewy_spadek + x_pk_lewy_2 - y_pk_lewy_2) / (1 + wsp_lewy_spadek), 12)
    y_pk_lewy_1 = y_pk_lewy_2 - x_pk_lewy_2 + x_pk_lewy_1

    if zalamanie_lewe.iloc[0, 0] == 'N':
        x_pk_lewy_4 = round((y_pk_lewy_3 + x_pk_lewy_3 - y_i_lewy - x_i_lewy * wsp_lewy_spadek) / (1 - wsp_lewy_spadek),
                            12)
        y_pk_lewy_4 = y_pk_lewy_3 - x_pk_lewy_4 + x_pk_lewy_3
    else:
        if zalamanie_lewe.iloc[0, 1] > 0:
            x_pk_lewy_4 = round(
                (y_pk_lewy_3 + x_pk_lewy_3 - y_i_lewy - x_i_lewy * wsp_lewy_spadek) / (1 - wsp_lewy_spadek),
                12)
            y_pk_lewy_4 = y_pk_lewy_3 - x_pk_lewy_4 + x_pk_lewy_3
        else:
            jez_lewa_spadek = round((i_lewy_1[1] - i_lewy_2[1]) / (i_lewy_2[0] - i_lewy_1[0]), 8)
            x_pk_lewy_4 = round(
                (y_pk_lewy_3 + x_pk_lewy_3 - y_i_lewy - x_i_lewy * jez_lewa_spadek) / (1 - jez_lewa_spadek),
                12)
            y_pk_lewy_4 = y_pk_lewy_3 - x_pk_lewy_4 + x_pk_lewy_3

    # Rysowanie linii:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_pk_lewy_1, y_pk_lewy_1, x_pk_lewy_2, y_pk_lewy_2))
    LWPline.Layer = 'AII_M_krawężnik'
    color.ColorIndex = 8
    LWPline.TrueColor = color
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_pk_lewy_3, y_pk_lewy_3, x_pk_lewy_4, y_pk_lewy_4))
    LWPline.Layer = 'AII_M_krawężnik'
    color.ColorIndex = 8
    LWPline.TrueColor = color

    # Krawężnik prawy:
    x_pk_prawy_2 = kraw_prawy[0]
    y_pk_prawy_2 = kraw_prawy[1]
    x_pk_prawy_3 = kraw_prawy[2]
    y_pk_prawy_3 = kraw_prawy[3]

    # Punkty podlewki:
    x_pk_prawy_4 = round(
        (y_pk_prawy_3 + x_pk_prawy_3 - y_i_prawy + x_i_prawy * wsp_prawy_spadek) / (1 + wsp_prawy_spadek),
        12)
    y_pk_prawy_4 = y_pk_prawy_3 + x_pk_prawy_3 - x_pk_prawy_4

    if zalamanie_prawe.iloc[0, 0] == 'N':
        x_pk_prawy_1 = round(
            (y_i_prawy - x_i_prawy * wsp_prawy_spadek + x_pk_prawy_2 - y_pk_prawy_2) / (1 - wsp_prawy_spadek), 12)
        y_pk_prawy_1 = y_pk_prawy_2 - x_pk_prawy_2 + x_pk_prawy_1
    else:
        if zalamanie_prawe.iloc[0, 1] > 0:
            x_pk_prawy_1 = round(
                (y_i_prawy - x_i_prawy * wsp_prawy_spadek + x_pk_prawy_2 - y_pk_prawy_2) / (1 - wsp_prawy_spadek), 12)
            y_pk_prawy_1 = y_pk_prawy_2 - x_pk_prawy_2 + x_pk_prawy_1
        else:
            jez_prawa_spadek = round((i_prawy_1[1] - i_prawy_2[1]) / (i_prawy_1[0] - i_prawy_2[0]), 8)
            x_pk_prawy_1 = round(
                (y_i_prawy - x_i_prawy * jez_prawa_spadek + x_pk_prawy_2 - y_pk_prawy_2) / (1 - jez_prawa_spadek), 12)
            y_pk_prawy_1 = y_pk_prawy_2 - x_pk_prawy_2 + x_pk_prawy_1

    # Rysowanie linii:
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_pk_prawy_1, y_pk_prawy_1, x_pk_prawy_2, y_pk_prawy_2))
    LWPline.Layer = 'AII_M_krawężnik'
    color.ColorIndex = 8
    LWPline.TrueColor = color
    LWPline = acad.model.AddLightWeightPolyline(aDouble(x_pk_prawy_3, y_pk_prawy_3, x_pk_prawy_4, y_pk_prawy_4))
    LWPline.Layer = 'AII_M_krawężnik'
    color.ColorIndex = 8
    LWPline.TrueColor = color

    print('PODLEWKI POD KRAWĘŻNIKI')

    # ==================================================================================================================
    # HATCH
    # ==================================================================================================================
    # Hatch = acad.model.AddHatch(1, 'ANSI33', True)
    # Obw_lewa = acad.model.AddLightWeightPolyline(aDouble(x_ch_lewy_2, y_ch_lewy_2, x_ch_lewy_1, y_ch_lewy_1,
    #                                                      x_pk_lewy_2, y_pk_lewy_2, x_pk_lewy_1, y_pk_lewy_1,
    #                                                      x_i_lewy_2, y_i_lewy_2))
    # Obw_lewa.Closed = True
    # Obw_prawa = acad.model.AddLightWeightPolyline(aDouble(x_ch_prawy_1, y_ch_prawy_1, x_ch_prawy_2, y_ch_prawy_2,
    #                                                       x_i_prawy_2, y_i_prawy_2, x_pk_prawy_4, y_pk_prawy_4,
    #                                                       x_pk_prawy_3, y_pk_prawy_3))
    # Obw_prawa.Closed = True
    # Hatch.AppendOuterLoop(Obw_prawa)
    # Hatch.AppendOuterLoop(Obw_lewa)
    # Hatch.Evaluate()

    # ==================================================================================================================
    # WYMIARY
    # ==================================================================================================================

    # Obliczenie położenia wymiarów:
    if bar_lewa == 'N':
        y_wym_l = y_bar_lewa + float(bar_lewa_rodz[-3:]) + 0.19
    else:
        if bal_lewa_rodz[0] == 'b':
            y_wym_l = y_bal_lewa + bal_lewa_wys + 0.165
        else:
            y_wym_l = y_bal_lewa + 1.24501667
    if bar_prawa == 'N':
        y_wym_p = y_bar_prawa + float(bar_prawa_rodz[-3:]) + 0.19
    else:
        if bal_prawa_rodz[0] == 'b':
            y_wym_p = y_bal_prawa + bal_prawa_wys + 0.165
        else:
            y_wym_p = y_bal_prawa + 1.24501667
    y_wym = y_g + max(1.75, y_wym_l, y_wym_p)

    # Wymiary szczegółowe:
    wymiary_gora_1 = []

    # Wymiary elementów jezdni po lewej stronie niwelety:
    # for index, row in pasy_lewe.iterrows():
    #     if index == 0:
    #         x1 = x_g
    #         y1 = y_g
    #     if isnan(pasy_lewe.iloc[index, 0]) or pasy_lewe.iloc[index, 0] == 0:
    #         break
        # x2 = round(x1 - row['PL - szer'], 8)
        # y2 = round(y1 + row['PL - szer'] * row['PL - spadek'] / 100, 8)
        # wymiary_gora_1.append((aDouble(x2, y2, 0), aDouble(x1, y1, 0), 'pas ruchu'))
        # x1 = x2
        # y1 = y2

    x1 = x_g
    y1 = y_g

    for width in pasy_lewe['PL - szer'].split('+'):            
        x2 = round(x1 - float(width)], 8)
        y2 = round(y1 + float(width)] * pasy_lewe['PL - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x2, y2, 0), aDouble(x1, y1, 0), 'pas ruchu'))

    for index, row in awaryjny_lewy.iterrows():
        if isnan(awaryjny_lewy.iloc[index, 0]) or awaryjny_lewy.iloc[index, 0] == 0:
            break
        x2 = round(x1 - row['PAL - szer'], 8)
        y2 = round(y1 + row['PAL - szer'] * row['PAL - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x2, y2, 0), aDouble(x1, y1, 0), 'pas awaryjny'))
        x1 = x2
        y1 = y2

    for index, row in opaska_lewa.iterrows():
        if isnan(opaska_lewa.iloc[index, 0]) or opaska_lewa.iloc[index, 0] == 0:
            break
        x2 = round(x1 - row['OL - szer'], 8)
        y2 = round(y1 + row['OL - szer'] * row['OL - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x2, y2, 0), aDouble(x1, y1, 0), 'opaska'))
        x1 = x2
        y1 = y2

    # Wymiary elementów jezdni po prawej stronie niwelety:
    for index, row in pasy_prawe.iterrows():
        if index == 0:
            x1 = x_g
            y1 = y_g
        if isnan(pasy_prawe.iloc[index, 0]) or pasy_prawe.iloc[index, 0] == 0:
            break
        x2 = round(x1 + row['PP - szer'], 8)
        y2 = round(y1 + row['PP - szer'] * row['PP - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x1, y1, 0), aDouble(x2, y2, 0), 'pas ruchu'))
        x1 = x2
        y1 = y2

    for index, row in awaryjny_prawy.iterrows():
        if isnan(awaryjny_prawy.iloc[index, 0]) or awaryjny_prawy.iloc[index, 0] == 0:
            break
        x2 = round(x1 + row['PAP - szer'], 8)
        y2 = round(y1 + row['PAP - szer'] * row['PAP - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x1, y1, 0), aDouble(x2, y2, 0), 'pas awaryjny'))
        x1 = x2
        y1 = y2

    for index, row in opaska_prawa.iterrows():
        if isnan(opaska_prawa.iloc[index, 0]) or opaska_prawa.iloc[index, 0] == 0:
            break
        x2 = round(x1 + row['OP - szer'], 8)
        y2 = round(y1 + row['OP - szer'] * row['OP - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x1, y1, 0), aDouble(x2, y2, 0), 'opaska'))
        x1 = x2
        y1 = y2

    # Wyrównanie wymiarów do góry krawężnika:
    wymiary_gora_1 = sorted(wymiary_gora_1)

    wymiary_gora_1[0][0][1] += 0.04
    wymiary_gora_1[-1][1][1] += 0.04

    # Usunięcie dociągnięcia wymiarów do jezdni:
    wymiary_gora_1[0][1][1] = 1000
    wymiary_gora_1[-1][0][1] = 1000
    for i in range(len(wymiary_gora_1) - 2):
        wymiary_gora_1[i + 1][0][1] = 1000
        wymiary_gora_1[i + 1][1][1] = 1000

    # Wymiary chodnika lewego:
    if bar_lewa == 'N':
        if ch_lewy_szer == 0:
            x_bar_wym = x_bar_lewa + 0.31
            y_bar_wym = round(y_bar_lewa + 0.76181926, 12)
            wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0), 0))
            wymiary_gora_1.append(
                (aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
        else:
            if ch_lewy_szer_cpr == 0:
                if ch_lewy_szer_sr == 0:
                    x_bar_wym = x_bar_lewa + 0.31
                    y_bar_wym = round(y_bar_lewa + 0.76181926, 12)
                    if ch_lewy_szer_ch >= 1.5:
                        wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0),
                                               'chodnik'))
                    else:
                        wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0),
                                               'przejście\Pdla obsługi'))
                    wymiary_gora_1.append(
                        (aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
                elif ch_lewy_szer_ch == 0:
                    x_bar_wym = x_bar_lewa + 0.31
                    y_bar_wym = round(y_bar_lewa + 0.76181926, 12)
                    wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0),
                                           'ścieżka rowerowa'))
                    wymiary_gora_1.append(
                        (aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
                else:
                    x_gra_wym = x_kraw_l - ch_lewy_szer_sr
                    # Dociągnięcie wymiaru granicy:
                    # y_gra_wym = y_kraw_l + 0.14 + (ch_lewy_szer_sr - 0.2) * ch_lewy_sp / 100
                    y_gra_wym = 1000

                    x_bar_wym = x_bar_lewa + 0.31
                    y_bar_wym = round(y_bar_lewa + 0.76181926, 12)
                    wymiary_gora_1.append((aDouble(x_gra_wym, y_gra_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0),
                                           'ścieżka rowerowa'))
                    wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_gra_wym, y_gra_wym, 0),
                                           'chodnik'))
                    wymiary_gora_1.append(
                        (aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
            else:
                x_bar_wym = x_bar_lewa + 0.31
                y_bar_wym = round(y_bar_lewa + 0.76181926, 12)
                wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0),
                                       'ciąg pieszo-rowerowy'))
                wymiary_gora_1.append(
                    (aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
    else:
        if bal_lewa_rodz[0] == 'b':
            x_bal_wym = x_bal_lewa + 0.04
            y_bal_wym = y_bal_lewa + bal_lewa_wys - 0.01
        else:
            x_bal_wym = x_bal_lewa + 0.175
            y_bal_wym = round(y_bal_lewa + 1.07001667, 12)
        if bar_lewa_rodz == 'bariera linowa':
            x_bar_wym = x_bar_lewa + 0.075
            y_bar_wym = round(y_bar_lewa + 0.72, 8)
            x_gra_wym_1 = x_bar_wym - bar_lewa_szer
            y_gra_wym_1 = y_bar_wym
        else:
            x_bar_wym = x_bar_lewa + 0.25
            y_bar_wym = round(y_bar_lewa + 0.73182587, 8)
            x_gra_wym_1 = x_bar_wym - bar_lewa_szer
            y_gra_wym_1 = 1000
        wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0), 0))
        wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
        if ch_lewy_szer_cpr == 0:
            if ch_lewy_szer_sr == 0:
                if ch_lewy_szer_ch >= 1.5:
                    wymiary_gora_1.append((aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0),
                                           'chodnik'))
                else:
                    wymiary_gora_1.append((aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0),
                                           'przejście\Pdla obsługi'))
            elif ch_lewy_szer_ch == 0:
                wymiary_gora_1.append((aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0),
                                       'ścieżka rowerowa'))
            else:
                x_gra_wym_2 = x_gra_wym_1 - ch_lewy_szer_sr
                # Dociągnięcie wymiaru granicy:
                # y_gra_wym_2 = round(y_kraw_l + 0.14 + (bar_lewa_op + bar_lewa_szer +
                #                                        ch_lewy_szer_sr - 0.2) * ch_lewy_sp / 100, 8)
                y_gra_wym_2 = 1000

                wymiary_gora_1.append((aDouble(x_gra_wym_2, y_gra_wym_2, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0),
                                       'ścieżka rowerowa'))
                wymiary_gora_1.append((aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_gra_wym_2, y_gra_wym_2, 0),
                                       'chodnik'))
        else:
            wymiary_gora_1.append((aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0),
                                   'ciąg pieszo-rowerowy'))

        wymiary_gora_1.append((aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_bal_wym, y_bal_wym, 0), 0))

    # Wymiary chodnika prawego:
    if bar_prawa == 'N':
        if ch_prawy_szer == 0:
            x_bar_wym = x_bar_prawa - 0.31
            y_bar_wym = round(y_bar_prawa + 0.76181926, 12)
            wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
            wymiary_gora_1.append(
                (aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0),
                 0))
        else:
            if ch_prawy_szer_cpr == 0:
                if ch_prawy_szer_sr == 0:
                    x_bar_wym = x_bar_prawa - 0.31
                    y_bar_wym = round(y_bar_prawa + 0.76181926, 12)
                    if ch_prawy_szer_ch >= 1.5:
                        wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0),
                                               'chodnik'))
                    else:
                        wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0),
                                               'przejście\Pdla obsługi'))
                    wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0),
                                           aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))
                elif ch_prawy_szer_ch == 0:
                    x_bar_wym = x_bar_prawa - 0.31
                    y_bar_wym = round(y_bar_prawa + 0.76181926, 12)
                    wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0),
                                           'ścieżka rowerowa'))
                    wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0),
                                           aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))
                else:
                    x_gra_wym = x_kraw_p + ch_prawy_szer_sr
                    # Dociągnięcie wymiaru granicy:
                    # y_gra_wym = y_kraw_p + 0.14 + (ch_prawy_szer_sr - 0.2) * ch_prawy_sp / 100
                    y_gra_wym = 1000

                    x_bar_wym = x_bar_prawa - 0.31
                    y_bar_wym = round(y_bar_prawa + 0.76181926, 12)
                    wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_gra_wym, y_gra_wym, 0),
                                           'ścieżka rowerowa'))
                    wymiary_gora_1.append((aDouble(x_gra_wym, y_gra_wym, 0), aDouble(x_bar_wym, y_bar_wym, 0),
                                           'chodnik'))
                    wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0),
                                           aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))
            else:
                x_bar_wym = x_bar_prawa - 0.31
                y_bar_wym = round(y_bar_prawa + 0.76181926, 12)
                wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0),
                                       'ciąg pieszo-rowerowy'))
                wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0),
                                       aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))
    else:
        if bal_prawa_rodz[0] == 'b':
            x_bal_wym = x_bal_prawa - 0.04
            y_bal_wym = y_bal_prawa + bal_prawa_wys - 0.01
        else:
            x_bal_wym = x_bal_prawa - 0.175
            y_bal_wym = round(y_bal_prawa + 1.07001667, 12)
        if bar_prawa_rodz == 'bariera linowa':
            x_bar_wym = x_bar_prawa - 0.075
            y_bar_wym = round(y_bar_prawa + 0.72, 8)
            x_gra_wym_1 = x_bar_wym + bar_prawa_szer
            y_gra_wym_1 = y_bar_wym
        else:
            x_bar_wym = x_bar_prawa - 0.25
            y_bar_wym = round(y_bar_prawa + 0.73182587, 8)
            x_gra_wym_1 = x_bar_wym + bar_prawa_szer
            y_gra_wym_1 = 1000
        wymiary_gora_1.append((aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_bar_wym, y_bar_wym, 0), 0))
        wymiary_gora_1.append((aDouble(x_bar_wym, y_bar_wym, 0), aDouble(x_gra_wym_1, y_gra_wym_1, 0), 0))
        if ch_prawy_szer_cpr == 0:
            if ch_prawy_szer_sr == 0:
                if ch_prawy_szer_ch >= 1.5:
                    wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_bal_wym, y_bal_wym, 0),
                                           'chodnik'))
                else:
                    wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_bal_wym, y_bal_wym, 0),
                                           'przejście\Pdla obsługi'))
            elif ch_prawy_szer_ch == 0:
                wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_bal_wym, y_bal_wym, 0),
                                       'ścieżka rowerowa'))
            else:
                x_gra_wym_2 = x_gra_wym_1 + ch_prawy_szer_sr
                # Dociągnięcie wymiaru granicy:
                # y_gra_wym_2 = round(y_kraw_p + 0.14 + (bar_prawa_op + bar_prawa_szer +
                #                                        ch_prawy_szer_sr - 0.2) * ch_prawy_sp / 100, 8)
                y_gra_wym_2 = 1000

                wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_gra_wym_2, y_gra_wym_2, 0),
                                       'ścieżka rowerowa'))
                wymiary_gora_1.append((aDouble(x_gra_wym_2, y_gra_wym_2, 0), aDouble(x_bal_wym, y_bal_wym, 0),
                                       'chodnik'))
        else:
            wymiary_gora_1.append((aDouble(x_gra_wym_1, y_gra_wym_1, 0), aDouble(x_bal_wym, y_bal_wym, 0),
                                   'ciąg pieszo-rowerowy'))

        wymiary_gora_1.append(
            (aDouble(x_bal_wym, y_bal_wym, 0), aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))

    # Wymiary główniejsze:
    wymiary_gora_2 = [(aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), aDouble(x_kraw_l, y_kraw_l + 0.04, 0), 0),
                      (aDouble(x_kraw_l, y_kraw_l + 0.04, 0), aDouble(x_kraw_p, y_kraw_p + 0.04, 0), 0),
                      (aDouble(x_kraw_p, y_kraw_p + 0.04, 0), aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0)]
    wymiary_gora_3 = [(aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0),
                       aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0),
                       0)]
    wymiary_gora_4 = []
    if ch_lewy_lat == 'T' and ch_prawy_lat == 'T':
        wymiary_gora_3.append((aDouble(x_ch_lewy_2 - 0.59, y_ch_lewy_2 - 0.04 + 0.55 * ch_lewy_sp / 100, 0),
                               aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), 0))
        wymiary_gora_3.append((aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0),
                               aDouble(x_ch_prawy_2 + 0.59, y_ch_prawy_2 - 0.04 + 0.55 * ch_prawy_sp / 100, 0), 0))
        wymiary_gora_4.append((aDouble(x_ch_lewy_2 - 0.59, y_ch_lewy_2 - 0.04 + 0.55 * ch_lewy_sp / 100, 0),
                               aDouble(x_ch_prawy_2 + 0.59, y_ch_prawy_2 - 0.04 + 0.55 * ch_prawy_sp / 100, 0), 0))
    if ch_lewy_lat == 'T' and ch_prawy_lat == 'N':
        wymiary_gora_3.append((aDouble(x_ch_lewy_2 - 0.59, y_ch_lewy_2 - 0.04 + 0.55 * ch_lewy_sp / 100, 0),
                               aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0), 0))
        wymiary_gora_4.append((aDouble(x_ch_lewy_2 - 0.59, y_ch_lewy_2 - 0.04 + 0.55 * ch_lewy_sp / 100, 0),
                               aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0), 0))
    if ch_lewy_lat == 'N' and ch_prawy_lat == 'T':
        wymiary_gora_3.append((aDouble(x_ch_prawy_2 + 0.04, y_ch_prawy_2 - 0.04, 0),
                               aDouble(x_ch_prawy_2 + 0.59, y_ch_prawy_2 - 0.04 + 0.55 * ch_prawy_sp / 100, 0), 0))
        wymiary_gora_4.append((aDouble(x_ch_lewy_2 - 0.04, y_ch_lewy_2 - 0.04, 0),
                               aDouble(x_ch_prawy_2 + 0.59, y_ch_prawy_2 - 0.04 + 0.55 * ch_prawy_sp / 100, 0), 0))

    # Wymiary boczne:
    if ch_lewy_lat == 'N':
        wym_d_lewa_1 = (aDouble(x_ch_lewy_2, y_ch_lewy_2, 0), aDouble(x_ch_lewy_2, y_ch_lewy_2 - ch_lewy_deska, 0), 90)
        wym_d_lewa_2 = aDouble(x_ch_lewy_2 - 0.415, y_ch_lewy_2, 0)
    else:
        wym_d_lewa_1 = (aDouble(x_ch_lewy_2 - 0.55, y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100, 0),
                        aDouble(x_ch_lewy_2 - 0.55, y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100 - ch_lewy_deska, 0), 90)
        wym_d_lewa_2 = aDouble(x_ch_lewy_2 - 0.725, y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100, 0)
    if ch_prawy_lat == 'N':
        wym_d_prawa_1 = (aDouble(x_ch_prawy_2, y_ch_prawy_2, 0),
                         aDouble(x_ch_prawy_2, y_ch_prawy_2 - ch_prawy_deska, 0), 90)
        wym_d_prawa_2 = aDouble(x_ch_prawy_2 + 0.55, y_ch_prawy_2, 0)
    else:
        wym_d_prawa_1 = (aDouble(x_ch_prawy_2 + 0.55, y_ch_prawy_2 + 0.55 * ch_prawy_sp / 100, 0),
                         aDouble(x_ch_prawy_2 + 0.55, y_ch_prawy_2 + 0.55 * ch_prawy_sp / 100 - ch_prawy_deska, 0), 90)
        wym_d_prawa_2 = aDouble(x_ch_prawy_2 + 0.84, y_ch_prawy_2, 0)
    Wym = acad.model.AddDimRotated(wym_d_lewa_1[0], wym_d_lewa_1[1], wym_d_lewa_2, radians(wym_d_lewa_1[2]))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(wym_d_prawa_1[0], wym_d_prawa_1[1], wym_d_prawa_2, radians(wym_d_prawa_1[2]))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    # ==================================================================================================================
    # KIERUNKI
    # ==================================================================================================================
    kierunki = []

    # Kierunki i punkty wstawienia:
    for index, row in pasy_lewe.iterrows():
        if index == 0:
            x1 = x_g
        if isnan(pasy_lewe.iloc[index, 0]) or pasy_lewe.iloc[index, 0] == 0:
            break
        x2 = round(x1 - row['PL - szer'], 8)
        kierunki.append([win32_point((x1 + x2) / 2, y_wym - 0.75, 0), row['PL - kier rodz'], row['PL - kier']])
        x1 = x2
    for index, row in pasy_prawe.iterrows():
        if index == 0:
            x1 = x_g
        if isnan(pasy_prawe.iloc[index, 0]) or pasy_prawe.iloc[index, 0] == 0:
            break
        x2 = round(x1 + row['PP - szer'], 8)
        kierunki.append([win32_point((x1 + x2) / 2, y_wym - 0.75, 0), row['PP - kier rodz'], row['PP - kier']])
        x1 = x2

    for kierunek in kierunki:
        Block = acad_32.ActiveDocument.ModelSpace.InsertBlock(kierunek[0], f'Kierunek_{kierunek[1]}', 1, 1, 1, 0)
        atr = Block.GetAttributes()
        atr[0].TextString = kierunek[2].upper()
        Block.Layer = 'AII_M_opis'

    # ==================================================================================================================
    # SPADKI
    # ==================================================================================================================
    spadki = []

    # Spadki jezdni:
    for index, row in pasy_prawe.iterrows():
        if isnan(pasy_prawe.iloc[index, 0]):
            if index == 0:
                x1 = x_g
                y1 = y_g
                x2 = x_g
                y2 = y_g
                spadek = 0
                spadek_gr = 0
            break
        if index == 0:
            x1 = x_g
            y1 = y_g
            spadek = row['PP - spadek']
            spadek_gr = spadek
            x2 = x_g + row['PP - szer']
            y2 = y_g + row['PP - szer'] * spadek / 100
        else:
            if row['PP - spadek'] == spadek:
                x2 += row['PP - szer']
                y2 += row['PP - szer'] * spadek / 100
            else:
                spadki.append([x1, y1, x2, y2])
                spadek = row['PP - spadek']
                x1 = x2
                y1 = y2
                x2 = x1 + row['PP - szer']
                y2 = y1 + row['PP - szer'] * spadek / 100

    for index, row in awaryjny_prawy.iterrows():
        if isnan(awaryjny_prawy.iloc[index, 0]) or awaryjny_prawy.iloc[index, 0] == 0:
            break
        else:
            if row['PAP - spadek'] == spadek:
                x2 += row['PAP - szer']
                y2 += row['PAP - szer'] * spadek / 100
            else:
                if len(spadki) >= 1:
                    spadki.append([x1, y1, x2, y2])
                spadek = row['PAP - spadek']
                x1 = x2
                y1 = y2
                x2 = x1 + row['PAP - szer']
                y2 = y1 + row['PAP - szer'] * spadek / 100

    for index, row in opaska_prawa.iterrows():
        if isnan(opaska_prawa.iloc[index, 0]) or opaska_prawa.iloc[index, 0] == 0:
            break
        else:
            if row['OP - spadek'] == spadek:
                x2 += row['OP - szer']
                y2 += row['OP - szer'] * spadek / 100
            else:
                if len(spadki) >= 1:
                    spadki.append([x1, y1, x2, y2])
                spadek = row['OP - spadek']
                x1 = x2
                y1 = y2
                x2 = x1 + row['OP - szer']
                y2 = y1 + row['OP - szer'] * spadek / 100

    for index, row in pasy_lewe.iterrows():
        if isnan(pasy_lewe.iloc[index, 0]):
            break
        if index == 0:
            if x2 == x_g:
                spadek = row['PL - spadek']
                x1 = x_g - row['PL - szer']
                y1 = y_g + row['PL - szer'] * spadek / 100
            else:
                spadek = row['PL - spadek']
                if spadek == -spadek_gr:
                    if x1 == x_g:
                        x1 -= row['PL - szer']
                        y1 += row['PL - szer'] * spadek / 100
                    else:
                        spadki.append([x1, y1, x2, y2])
                        x2 = spadki[0][2]
                        y2 = spadki[0][3]
                        spadki.pop(0)
                        x1 = -row['PL - szer']
                        y1 = row['PL - szer'] * spadek / 100
                else:
                    spadki.append([x1, y1, x2, y2])
                    x2 = x_g
                    y2 = y_g
                    x1 = x_g - row['PL - szer']
                    y1 = y_g + row['PL - szer'] * spadek / 100
        else:
            if row['PL - spadek'] == spadek:
                x1 -= row['PL - szer']
                y1 += row['PL - szer'] * spadek / 100
            else:
                spadki.append([x1, y1, x2, y2])
                spadek = row['PL - spadek']
                x2 = x1
                y2 = y1
                x1 = x2 - row['PL - szer']
                y1 = y2 + row['PL - szer'] * spadek / 100

    for index, row in awaryjny_lewy.iterrows():
        if isnan(awaryjny_lewy.iloc[index, 0]) or awaryjny_lewy.iloc[index, 0] == 0:
            break
        else:
            if x1 == x_g and row['PAL - spadek'] == -spadek_gr:
                spadek = row['PAL - spadek']
                x1 -= row['PAL - szer']
                y1 += row['PAL - szer'] * spadek / 100
            elif row['PAL - spadek'] == spadek:
                x1 -= row['PAL - szer']
                y1 += row['PAL - szer'] * spadek / 100
            else:
                spadki.append([x1, y1, x2, y2])
                spadek = row['PAL - spadek']
                x2 = x1
                y2 = y1
                x1 = x2 - row['PAL - szer']
                y1 = y2 + row['PAL - szer'] * spadek / 100

    for index, row in opaska_lewa.iterrows():
        if isnan(opaska_lewa.iloc[index, 0]) or opaska_lewa.iloc[index, 0] == 0:
            break
        else:
            if x1 == x_g and row['OL - spadek'] == -spadek_gr:
                spadek = row['OL - spadek']
                x1 -= row['OL - szer']
                y1 += row['OL - szer'] * spadek / 100
            elif row['OL - spadek'] == spadek:
                x1 -= row['OL - szer']
                y1 += row['OL - szer'] * spadek / 100
            else:
                spadki.append([x1, y1, x2, y2])
                spadek = row['OL - spadek']
                x2 = x1
                y2 = y1
                x1 = x2 - row['OL - szer']
                y1 = y2 + row['OL - szer'] * spadek / 100
    spadki.append([x1, y1, x2, y2])

    # Spadki chodników:
    if bar_lewa == 'N':
        x2 = x_ch_lewy_1
        y2 = y_ch_lewy_1
        x1 = x_bar_lewa + 0.31
        y1 = y2 + (x2 - x1) * ch_lewy_sp / 100
    else:
        x2 = x_kraw_l - bar_lewa_op - bar_lewa_szer
        y2 = round(y_kraw_l + 0.14 + (bar_lewa_op + bar_lewa_szer - 0.2) * ch_lewy_sp / 100, 8)
        x1 = x2 - ch_lewy_szer
        y1 = y2 + ch_lewy_szer * ch_lewy_sp / 100
    if x1 == x2:
        x1 = x2 - 0.01
        y1 = round(y2 + 0.01 * ch_lewy_sp / 100, 8)
    spadki.append([x1, y1, x2, y2])

    if bar_prawa == 'N':
        x1 = x_ch_prawy_1
        y1 = y_ch_prawy_1
        x2 = x_bar_prawa - 0.31
        y2 = y1 + (x2 - x1) * ch_prawy_sp / 100
    else:
        x1 = x_kraw_p + bar_prawa_op + bar_prawa_szer
        y1 = round(y_kraw_p + 0.14 + (bar_prawa_op + bar_prawa_szer - 0.2) * ch_prawy_sp / 100, 8)
        x2 = x1 + ch_prawy_szer
        y2 = y1 + ch_prawy_szer * ch_prawy_sp / 100
    if x1 == x2:
        x2 = x1 + 0.01
        y2 = round(y1 + 0.01 * ch_prawy_sp / 100, 8)
    spadki.append([x1, y1, x2, y2])

    for spadek in spadki:
        x = round((spadek[0] + spadek[2]) / 2, 8)
        y = round((spadek[1] + spadek[3]) / 2 + 0.05, 8)
        ins_point = win32_point(x, y, 0)
        angle = -atan(fabs((spadek[3] - spadek[1]) / (spadek[2] - spadek[0])))
        spadek_wartosc = round(fabs((spadek[3] - spadek[1]) / (spadek[2] - spadek[0]) * 100), 1)
        Block = acad_32.ActiveDocument.ModelSpace.InsertBlock(ins_point, 'Spadek', 1, 1, 1, angle)
        atr = Block.GetAttributes()
        atr[0].TextString = f'{spadek_wartosc}%'
        if spadek[3] > spadek[1]:
            Block2 = Block.Mirror(win32_point(x, y, 0), win32_point(x, y + 1, 0))
            Block.Delete()
            Block = Block2
        Block.Layer = 'AII_M_wymiary'

    # ==================================================================================================================
    # KOTY WYSOKOŚCIOWE
    # ==================================================================================================================
    koty = [[x_kraw_l, y_kraw_l, 0, 1], [x_kraw_p, y_kraw_p, 0, -1], [x_kraw_l - 0.20, y_kraw_l + 0.14, 0, -1],
            [x_kraw_p + 0.20, y_kraw_p + 0.14, 0, 1], [x_ch_prawy_2, y_ch_prawy_2, 0, 1],
            [x_ch_lewy_2, y_ch_lewy_2, 0, -1]]

    if x_kraw_l != x_g and x_kraw_p != x_g:
        koty.append([x_g, y_g, 0, 1])

    if ch_lewy_lat == 'T':
        koty.append([x_ch_lewy_2 - 0.55, y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100, 0, -1])
    if ch_prawy_lat == 'T':
        koty.append([x_ch_prawy_2 + 0.55, y_ch_prawy_2 + 0.55 * ch_prawy_sp / 100, 0, 1])

    for index, kota in enumerate(koty):
        if kota[1] == y_g:
            tekst = f'%%p{kota[1] - y_g:.3f}'
        elif kota[1] > y_g:
            tekst = f'+{kota[1] - y_g:.3f}'
        else:
            tekst = f'{kota[1] - y_g:.3f}'
        koty[index] = [kota[0], kota[1], kota[2], kota[3], tekst]

    # ==================================================================================================================
    # OŚ NIWELETY
    # ==================================================================================================================
    os = [x_g, y_g - 0.5, x_g]
    Text = acad.model.AddText('oś niwelety', aDouble(x_g - 0.075, y_g + 0.3, 0), 0.125)
    Text.Rotation = radians(90)
    Text.StyleName = 'AII_norm.'
    Text.Layer = 'AII_M_opis'

    # ==================================================================================================================
    # USTRÓJ NOŚNY
    # ==================================================================================================================

    # Górna powierzchnia konstrukcji:
    pow_gorna = [[i[0], round(i[1] - 0.01, 8)] for i in jezdnia_3]
    pl_gor_lewy_1 = pow_gorna[0]
    pl_gor_lewy_2 = pow_gorna[1]
    pl_gor_prawy_1 = pow_gorna[-2]
    pl_gor_prawy_2 = pow_gorna[-1]
    pow_gorna = pow_gorna[1:-1]
    x_pl_gor_lewy = pl_gor_lewy_1[0] + 0.01
    x_pl_gor_prawy = pl_gor_prawy_2[0] - 0.01
    y_pl_gor_lewy = round(pl_gor_lewy_1[1] + (pl_gor_lewy_2[1] - pl_gor_lewy_1[1]) *
                            (x_pl_gor_lewy - pl_gor_lewy_1[0]) / (pl_gor_lewy_2[0] - pl_gor_lewy_1[0]), 8)
    y_pl_gor_prawy = round(pl_gor_prawy_1[1] + (pl_gor_prawy_2[1] - pl_gor_prawy_1[1]) *
                            (x_pl_gor_prawy - pl_gor_prawy_1[0]) / (pl_gor_prawy_2[0] - pl_gor_prawy_1[0]), 8)

    pow_gorna.insert(0,[x_pl_gor_lewy, y_pl_gor_lewy])
    pow_gorna.append([x_pl_gor_prawy, y_pl_gor_prawy])
    pow_gorna = [[round(i[0], 8), round(i[1], 8)] for i in pow_gorna]

    # Usuwanie punktów pośrednich:
    pow_gorna_del = []
    for i in range(len(pow_gorna) - 2):
        xi = pow_gorna[i + 1][0]
        yi = pow_gorna[i + 1][1]
        x1 = pow_gorna[i][0]
        y1 = pow_gorna[i][1]
        x2 = pow_gorna[i + 2][0]
        y2 = pow_gorna[i + 2][1]
        if round((y2 - yi) / (x2 - xi), 6) == round((yi - y1) / (xi - x1), 6):
            pow_gorna_del.append([xi, yi])

    for i in pow_gorna_del:
        pow_gorna.remove(i)

    # ==================================================================================================================
    # PUNKTY SKRAJNE
    # ==================================================================================================================
    if ch_lewy_lat == 'N' and ch_prawy_lat == 'N':
        x_lewy_skr = x_ch_lewy_2 - 0.04
        y_lewy_skr = y_ch_lewy_2 - 0.04
        x_prawy_skr = x_ch_prawy_2 + 0.04
        y_prawy_skr = y_ch_prawy_2 - 0.04
    elif ch_lewy_lat == 'T' and ch_prawy_lat == 'N':
        x_lewy_skr = x_ch_lewy_2 - 0.59
        y_lewy_skr = y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100 - 0.04
        x_prawy_skr = x_ch_prawy_2 + 0.04
        y_prawy_skr = y_ch_prawy_2 - 0.04
    elif ch_lewy_lat == 'N' and ch_prawy_lat == 'T':
        x_lewy_skr = x_ch_lewy_2 - 0.04
        y_lewy_skr = y_ch_lewy_2 - 0.04
        x_prawy_skr = x_ch_prawy_2 + 0.59
        y_prawy_skr = y_ch_prawy_2 + 0.55 * ch_prawy_sp / 100 - 0.04
    else:
        x_lewy_skr = x_ch_lewy_2 - 0.59
        y_lewy_skr = y_ch_lewy_2 + 0.55 * ch_lewy_sp / 100 - 0.04
        x_prawy_skr = x_ch_prawy_2 + 0.59
        y_prawy_skr = y_ch_prawy_2 + 0.55 * ch_prawy_sp / 100 - 0.04
    print('WYMIARY')



    print(f'Obiekt {obiekt} - przekrój uzytkowy narysowany!')

    return [y_wym, wymiary_gora_1, wymiary_gora_2, wymiary_gora_3, wymiary_gora_4,
            x_lewy_skr, y_lewy_skr, x_prawy_skr, y_prawy_skr, koty, pow_gorna, os]
