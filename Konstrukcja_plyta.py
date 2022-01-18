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
def rysowanie_konstrukcja_plytowy(pow_gorna):
    # Pobieranie danych:

    obiekt=Pobieranie_danych.obiekt
    konstrukcja = Pobieranie_danych.konstrukcja   

    h_wsp_kr = konstrukcja['P_WSP - h'][0]
    h_wsp_zam = konstrukcja['P_WSP - h zam'][0]
    h_pl_pl = konstrukcja['P_PL - płaski spód'][0]
    h_pl = konstrukcja['P_PL - h'][0]
    b_pl = konstrukcja['P_PL - b'][0]
    b_pl_skos_l = konstrukcja['P_PL - b skos L'][0]
    b_pl_skos_p = konstrukcja['P_PL - b skos P'][0]

    # Dolna powierzchnia konstrukcji:
    x_kon = round((pow_gorna[-1][0] + pow_gorna[0][0]) / 2, 8)
    for i in range(len(pow_gorna)):
        x2 = pow_gorna[i][0]
        if x_kon <= x2:
            index = i
            break
    x1 = pow_gorna[index - 1][0]
    y1 = pow_gorna[index - 1][1]
    x2 = pow_gorna[index][0]
    y2 = pow_gorna[index][1]
    y_kon = y1 + (y2 - y1) * (x_kon - x1) / (x2 - x1)

    x_3 = round(x_kon - b_pl / 2, 8)
    x_4 = round(x_kon + b_pl / 2, 8)
    x_2 = round(x_3 - b_pl_skos_l, 8)
    x_5 = round(x_4 + b_pl_skos_p, 8)
    x_1 = pow_gorna[0][0]
    x_6 = pow_gorna[-1][0]

    pow_dolna_x = []
    pow_dolna_x.append(x_1)
    if h_wsp_zam != h_pl:
        pow_dolna_x.append(x_2)
    pow_dolna_x.append(x_3)
    pow_dolna_x.append(x_4)
    if h_wsp_zam != h_pl:
        pow_dolna_x.append(x_5)
    pow_dolna_x.append(x_6)

    # Obliczenie współrzędnych y dla punktów charakterystycznych:
    pow_dolna = []

    for x in pow_dolna_x:
        for i in range(len(pow_gorna)):
            x2 = pow_gorna[i][0]
            if x <= x2:
                index = i
                break
        x1 = pow_gorna[index - 1][0]
        y1 = pow_gorna[index - 1][1]
        x2 = pow_gorna[index][0]
        y2 = pow_gorna[index][1]
        y = y1 + (y2 - y1) * (x - x1) / (x2 - x1)
        pow_dolna.append([round(x, 6), round(y, 6)])

    pow_dolna = sorted(pow_dolna)

    for i, pkt in enumerate(pow_dolna):
        if len(pow_dolna) == 6:
            if i in [0, 5]:
                y = round(pkt[1] - h_wsp_kr, 6)
            elif i in [1, 4]:
                y = round(pkt[1] - h_wsp_zam, 6)
            elif i in [2, 3]:
                y = round(pkt[1] - h_pl, 6)
        else:
            if i in [0, 3]:
                y = round(pkt[1] - h_wsp_kr, 6)
            elif i in [1, 2]:
                y = round(pkt[1] - h_pl, 6)
        pkt[1] = y

    if h_pl_pl == 'T':
        pow_gorna_pl = [i for i in pow_gorna if x_3 <= i[0] <= x_4]
        pow_gorna_min = sorted(pow_gorna_pl, key=lambda i: i[1])
        x_min_st = pow_gorna_min[0][0]
        y_min_st = pow_gorna_min[0][1]
        y_min_st = round(y_min_st - h_pl, 8)
        pow_dolna[2][1] = y_min_st
        pow_dolna[-3][1] = y_min_st

    y_1 = pow_dolna[0][1]
    y_6 = pow_dolna[-1][1]

    if len(pow_dolna) == 6:
        y_2 = pow_dolna[1][1]
        y_3 = pow_dolna[2][1]
        y_4 = pow_dolna[3][1]
        y_5 = pow_dolna[4][1]
    else:
        y_3 = pow_dolna[1][1]
        y_4 = pow_dolna[2][1]

    konstrukcja_lista = list(chain.from_iterable(pow_gorna + pow_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True

    # ==================================================================================================================
    # OSIE I OPISY
    # ==================================================================================================================

    # Rzędna minimalna konstrukcji:
    y_min = min([i[1] for i in pow_dolna])
    y_dol = round((pow_dolna[2][1] + pow_dolna[-3][1]) / 2, 8)
    LWPline = acad.model.AddLightWeightPolyline(aDouble([x_kon, y_min - 0.325, x_kon, y_kon + 0.2]))
    LWPline.LinetypeScale = 0.1
    LWPline.Layer = 'AII_M_osie główne'
    x_text = x_kon - 0.075
    y_text = (y_dol + y_kon) / 2
    text = 'oś konstrukcji'
    Text = acad.model.AddText(text, aDouble(x_text, y_text, 0), 0.125)
    Text.Alignment = 1
    Text.TextAlignmentPoint = aDouble(x_text, y_text, 0)
    Text.Rotation = radians(90)
    Text.StyleName = 'AII_norm.'
    Text.Layer = 'AII_M_opis'

    # ==================================================================================================================
    # WYMIARY DÓŁ
    # ==================================================================================================================

    # Wymiary poziome:

    # Poziomy wymiarów:
    y_wym_1 = aDouble(0, y_min - 0.25, 0)
    y_wym_2 = aDouble(0, y_min - 0.5, 0)
    y_wym_3 = aDouble(0, y_min - 0.75, 0)

    # Wymiary:
    wymiary_dol_1 = [[aDouble(x_3, y_3, 0), aDouble(x_kon, y_min - 0.325, 0), 0],
                     [aDouble(x_kon, y_min - 0.325, 0), aDouble(x_4, y_4, 0), 0]]
    wymiary_dol_2 = []

    if x_2 == x_3 or h_wsp_zam == h_pl:
        wymiary_dol_2.append([aDouble(x_1, y_1, 0), aDouble(x_3, y_3, 0), 0])
    else:
        wymiary_dol_2.append([aDouble(x_1, y_1, 0), aDouble(x_2, y_2, 0), 0])
        wymiary_dol_2.append([aDouble(x_2, y_2, 0), aDouble(x_3, y_3, 0), 0])

    wymiary_dol_2.append([aDouble(x_3, y_3, 0), aDouble(x_4, y_4, 0), 0])

    if x_4 == x_5 or h_wsp_zam == h_pl:
        wymiary_dol_2.append([aDouble(x_4, y_4, 0), aDouble(x_6, y_6, 0), 0])
    else:
        wymiary_dol_2.append([aDouble(x_4, y_4, 0), aDouble(x_5, y_5, 0), 0])
        wymiary_dol_2.append([aDouble(x_5, y_5, 0), aDouble(x_6, y_6, 0), 0])

    x_lewy_skr_dol = pow_dolna[0][0]
    y_lewy_skr_dol = pow_dolna[0][1]
    x_prawy_skr_dol = pow_dolna[-1][0]
    y_prawy_skr_dol = pow_dolna[-1][1]
    wymiary_dol_3 = [[aDouble(x_lewy_skr_dol, y_lewy_skr_dol, 0), aDouble(x_prawy_skr_dol, y_prawy_skr_dol, 0), 0]]

    # Wstawienie wymiarów:
    for index, wymiar in enumerate(wymiary_dol_1):
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_1, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        if index == 0:
            Wym.ExtLine2Suppress = True
        else:
            Wym.ExtLine1Suppress = True

    for index, wymiar in enumerate(wymiary_dol_2):
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_2, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for wymiar in wymiary_dol_3:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_3, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # Wymiary pionowe:

    # Deski:
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[0][0], pow_gorna[0][1] - h_wsp_kr, 0),
                                   aDouble(pow_gorna[0][0], pow_gorna[0][1], 0),
                                   aDouble(pow_gorna[0][0] - 0.175, pow_gorna[0][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[-1][0], pow_gorna[-1][1] - h_wsp_kr, 0),
                                   aDouble(pow_gorna[-1][0], pow_gorna[-1][1], 0),
                                   aDouble(pow_gorna[-1][0] + 0.3, pow_gorna[-1][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    # Zamocowania:
    if h_wsp_zam != h_pl:
        xy_zam = [[x_2, y_2], [x_5, y_5]]
        for index, pkt in enumerate(xy_zam):
            x = pkt[0]
            y1 = pkt[1]
            y2 = pkt[1] + h_wsp_zam
            if index == 0:
                x_wym = x + 0.175
            else:
                x_wym = x - 0.175
            Wym = acad.model.AddDimRotated(aDouble(x, y1, 0), aDouble(x, y2, 0), aDouble(x_wym, y1, 0), radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'

    # Płyty:
    if h_pl_pl == 'N':
        Wym = acad.model.AddDimRotated(aDouble(x_kon, y_dol, 0),
                                       aDouble(x_kon, y_kon, 0),
                                       aDouble(x_kon + 0.25, pkt[1], 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
    else:
        Wym = acad.model.AddDimRotated(aDouble(x_min_st, y_min_st, 0),
                                       aDouble(x_min_st, y_min_st + h_pl, 0),
                                       aDouble(x_min_st + 0.25, y_min_st, 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    print(f'Obiekt {obiekt} - ustrój betonowy, płytowy narysowany!')
