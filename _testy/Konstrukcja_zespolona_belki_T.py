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
def rysowanie_konstrukcja_belki_T(pow_gorna):
    # Pobranie danych:
    # konstrukcja = pd.read_excel(file, usecols=[43, 44], sheet_name=sheet)
    # h_pl = konstrukcja.iloc[0, 0]
    # bel_rodz = konstrukcja.iloc[0, 1]\

    obiekt=Pobieranie_danych.obiekt
    konstrukcja = Pobieranie_danych.konstrukcja   
    h_pl = konstrukcja['T_PL - h'][0]
    bel_rodz = konstrukcja['T_Belka T'][0] 

    # Dolna powierzchnia płyty:
    pow_dolna = [[i[0], round(i[1] - h_pl, 8)] for i in pow_gorna]

    plyta_lista = list(chain.from_iterable(pow_gorna + pow_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(plyta_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True

    # Wstawianie belek:
    x_bel = []

    # Spadek daszkowy:
    if len(pow_dolna) == 5:
        # Wspornik lewy:
        x1 = pow_dolna[0][0]
        x2 = pow_dolna[1][0]
        x = x2 - 0.45
        x_bel.append(x)
        del_x = round(x - x1, 8)
        while del_x >= 1.35:
            x = round(x - 0.9, 8)
            x_bel.append(x)
            del_x = round(x - x1, 8)

        # Daszek lewy:
        x1 = pow_dolna[1][0]
        x2 = pow_dolna[2][0]
        x = x2 - 0.45
        x_bel.append(x)
        del_x = round(x - x1, 8)
        while del_x >= 1.35:
            x = round(x - 0.9, 8)
            x_bel.append(x)
            del_x = round(x - x1, 8)

        # Daszek prawy:
        x1 = pow_dolna[2][0]
        x2 = pow_dolna[3][0]
        x = x1 + 0.45
        x_bel.append(x)
        del_x = round(x2 - x, 8)
        while del_x >= 1.35:
            x = round(x + 0.9, 8)
            x_bel.append(x)
            del_x = round(x2 - x, 8)

        # Wspornik prawy:
        x1 = pow_dolna[3][0]
        x2 = pow_dolna[4][0]
        x = x1 + 0.45
        x_bel.append(x)
        del_x = round(x2 - x, 8)
        while del_x >= 1.35:
            x = round(x + 0.9, 8)
            x_bel.append(x)
            del_x = round(x2 - x, 8)
    elif len(pow_dolna) == 3:
        # Część lewa:
        x1 = pow_dolna[0][0]
        x2 = pow_dolna[1][0]
        x = x2 - 0.45
        x_bel.append(x)
        del_x = round(x - x1, 8)
        while del_x >= 1.35:
            x = round(x - 0.9, 8)
            x_bel.append(x)
            del_x = round(x - x1, 8)

        # Część prawa:
        x1 = pow_dolna[1][0]
        x2 = pow_dolna[2][0]
        x = x1 + 0.45
        x_bel.append(x)
        del_x = round(x2 - x, 8)
        while del_x >= 1.35:
            x = round(x + 0.9, 8)
            x_bel.append(x)
            del_x = round(x2 - x, 8)

    x_bel = sorted(x_bel)

    # Obliczenie współrzędnych y i kątów dla punktów wstawienia belek:
    xy_bel = []
    for x in x_bel:
        for i in range(len(pow_dolna)):
            x2 = pow_dolna[i][0]
            if x <= x2:
                index = i
                break
        x1 = pow_dolna[index - 1][0]
        y1 = pow_dolna[index - 1][1]
        x2 = pow_dolna[index][0]
        y2 = pow_dolna[index][1]
        y = y1 + (y2 - y1) * (x - x1) / (x2 - x1)
        angle = atan((y2 - y1) / (x2 - x1))
        xy_bel.append([x, y, angle])

    # Wstawianie belek:
    for belka in xy_bel:
        Belka = acad.model.InsertBlock(aDouble(belka[0], belka[1], 0), f'Belka_{bel_rodz}', 1, 1, 1, belka[2])
        Belka.Layer = 'AII_M_konstrukcja beton'
    
    # Punkty skrajne:
    x_lewy_skr_dol = pow_dolna[0][0]
    y_lewy_skr_dol = pow_dolna[0][1]
    x_prawy_skr_dol = pow_dolna[-1][0]
    y_prawy_skr_dol = pow_dolna[-1][1]

    # ==================================================================================================================
    # WYMIARY DÓŁ
    # ==================================================================================================================
    delta_y = {'T12': 0.9, 'T15': 1.05, 'T18': 1.05, 'T21': 1.2, 'T24': 1.3, 'T27': 1.4}
    y_min = min([i[1] for i in pow_dolna])
    y_wym_d = y_min - delta_y[bel_rodz] + 0.175

    # Poziomy wymiarów:
    y_wym_1 = aDouble(0, y_min - delta_y[bel_rodz], 0)
    y_wym_2 = aDouble(0, y_min - delta_y[bel_rodz] - 0.25, 0)

    # Wymiary poziome:
    wymiary_dol_1 = []
    wymiary_dol_2 = [[aDouble(pow_dolna[0][0], pow_dolna[0][1], 0), aDouble(pow_dolna[-1][0], pow_dolna[-1][1], 0), 0]]

    for i, pkt in enumerate(xy_bel):
        if i == 0:
            wymiary_dol_1.append([aDouble(pow_dolna[0][0], pow_dolna[0][1], 0), aDouble(pkt[0], y_wym_d, 0), 0])
            wymiary_dol_1.append([aDouble(pkt[0], y_wym_d, 0), aDouble(xy_bel[i + 1][0], y_wym_d, 0)])
        elif i == len(xy_bel) - 1:
            wymiary_dol_1.append([aDouble(pkt[0], y_wym_d, 0), aDouble(pow_dolna[-1][0], pow_dolna[-1][1], 0)])
        else:
            wymiary_dol_1.append([aDouble(pkt[0], y_wym_d, 0), aDouble(xy_bel[i + 1][0], y_wym_d, 0)])

    # Wstawienie wymiarów:
    for wymiar in wymiary_dol_1:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_1, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for wymiar in wymiary_dol_2:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_2, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # Wymiary pionowe:
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[0][0], pow_gorna[0][1] - h_pl, 0),
                                   aDouble(pow_gorna[0][0], pow_gorna[0][1], 0),
                                   aDouble(pow_gorna[0][0] - 0.175, pow_gorna[0][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[-1][0], pow_gorna[-1][1] - h_pl, 0),
                                   aDouble(pow_gorna[-1][0], pow_gorna[-1][1], 0),
                                   aDouble(pow_gorna[-1][0] + 0.3, pow_gorna[-1][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'


    print(f'Obiekt {obiekt} - ustrój z belek T narysowany!')
