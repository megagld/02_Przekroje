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
def rysowanie_konstrukcja_skrzynkowy(y_g, pow_gorna):

    # Pobranie danych:
    # konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63], sheet_name=sheet)

    obiekt=Pobieranie_danych.obiekt
    konstrukcja = Pobieranie_danych.konstrukcja   

    h_wsp_kr = konstrukcja['S_WSP - h'][0]
    h_wsp_zam = konstrukcja['S_WSP - h zam'][0]
    h_skrz = konstrukcja['S_SKRZ - h'][0]
    b_skrz = konstrukcja['S_SKRZ - b'][0]
    b_skos_l = konstrukcja['S_SKRZ - skos L'][0]
    b_skos_p = konstrukcja['S_SKRZ - skos P'][0]
    h_plg = konstrukcja['S_PLG - h'][0]
    h_plg_zam_l = konstrukcja['S_PLG - h zam L'][0]
    b1_plg_zam_l = konstrukcja['S_PLG - b1 zam L'][0]
    b2_plg_zam_l = konstrukcja['S_PLG - b2 zam L'][0]
    h_plg_zam_p = konstrukcja['S_PLG - h zam P'][0]
    b1_plg_zam_p = konstrukcja['S_PLG - b1 zam P'][0]
    b2_plg_zam_p = konstrukcja['S_PLG - b2 zam P'][0]
    h_pld = konstrukcja['S_PLD - h'][0]
    h_pld_zam_l = konstrukcja['S_PLD - h zam L'][0]
    b1_pld_zam_l = konstrukcja['S_PLD - b1 zam L'][0]
    b2_pld_zam_l = konstrukcja['S_PLD - b2 zam L'][0]
    h_pld_zam_p = konstrukcja['S_PLD - h zam P'][0]
    b1_pld_zam_p = konstrukcja['S_PLD - b1 zam P'][0]
    b2_pld_zam_p = konstrukcja['S_PLD - b2 zam P'][0]
    t_sr = konstrukcja['S_T - gr'][0]

    # Dolna powierzchnia konstrukcji:
    x_kon = round((pow_gorna[-1][0] + pow_gorna[0][0]) / 2, 6)

    # Współrzędne x powierzchni dolnej:
    x_1 = pow_gorna[0][0]
    x_2 = round(x_kon - b_skrz / 2 - b_skos_l, 6)
    x_3 = round(x_kon - b_skrz / 2, 6)
    x_4 = round(x_kon + b_skrz / 2, 6)
    x_5 = round(x_kon + b_skrz / 2 + b_skos_p, 6)
    x_6 = pow_gorna[-1][0]
    x_dolna = [x_1, x_2, x_3, x_4, x_5, x_6]

    # Współrzędna y skrzynki:
    for i in range(len(pow_gorna)):
        x2 = pow_gorna[i][0]
        if x_kon <= x2:
            index = i
            break
    x1 = pow_gorna[index - 1][0]
    y1 = pow_gorna[index - 1][1]
    x2 = pow_gorna[index][0]
    y2 = pow_gorna[index][1]
    y_skrz = round(y1 + (y2 - y1) * (x_kon - x1) / (x2 - x1), 6)

    pow_dolna = []

    for x in x_dolna:
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

    for index, pkt in enumerate(pow_dolna):
        if index in [0, 5]:
            y = round(pkt[1] - h_wsp_kr, 6)
        elif index in [1, 4]:
            y = round(pkt[1] - h_wsp_zam, 6)
        elif index in [2, 3]:
            y = y_skrz - h_skrz
        pkt[1] = y

    x_1 = pow_dolna[0][0]
    y_1 = pow_dolna[0][1]
    x_2 = pow_dolna[1][0]
    y_2 = pow_dolna[1][1]
    x_3 = pow_dolna[2][0]
    y_3 = pow_dolna[2][1]
    x_4 = pow_dolna[3][0]
    y_4 = pow_dolna[3][1]
    x_5 = pow_dolna[4][0]
    y_5 = pow_dolna[4][1]
    x_6 = pow_dolna[5][0]
    y_6 = pow_dolna[5][1]

    # Dolna powierzchnia wewnętrzna:
    if b_skos_l != 0:
        t_x_sr_l = round(t_sr / sin(atan((y_2 - y_3) / (x_3 - x_2))), 12)
    else:
        t_x_sr_l = t_sr
    if b_skos_p != 0:
        t_x_sr_p = round(t_sr / sin(atan((y_5 - y_4) / (x_5 - x_4))), 12)
    else:
        t_x_sr_p = t_sr

    pow_wew_dolna = []
    if b1_pld_zam_l + b2_pld_zam_l != 0:
        y_7 = round(y_3 + h_pld_zam_l, 12)
    else:
        y_7 = round(y_3 + h_pld, 12)
    x_7 = round(t_x_sr_l + x_3 - (y_7 - y_3) * (x_3 - x_2) / (y_2 - y_3), 12)
    pow_wew_dolna.append([x_7, y_7])
    x_8 = round(x_7 + b1_pld_zam_l, 12)
    y_8 = y_7
    x_9 = round(x_8 + b2_pld_zam_l, 12)
    y_9 = round(y_3 + h_pld, 12)
    if b1_pld_zam_l != 0 and b2_pld_zam_l != 0:
        pow_wew_dolna.append([x_8, y_8])
        pow_wew_dolna.append([x_9, y_9])
    elif b1_pld_zam_l == 0 and b2_pld_zam_l != 0:
        pow_wew_dolna.append([x_9, y_9])
    elif b1_pld_zam_l != 0 and b2_pld_zam_l == 0:
        pow_wew_dolna.append([x_8, y_8])
        pow_wew_dolna.append([x_9, y_9])

    if b1_pld_zam_p + b2_pld_zam_p != 0:
        y_12 = round(y_4 + h_pld_zam_p, 12)
    else:
        y_12 = round(y_4 + h_pld, 12)

    x_12 = round(x_4 + (y_12 - y_4) * (x_5 - x_4) / (y_5 - y_4) - t_x_sr_p, 12)
    x_11 = round(x_12 - b1_pld_zam_p, 12)
    y_11 = y_12
    x_10 = round(x_11 - b2_pld_zam_p, 12)
    y_10 = round(y_4 + h_pld, 12)
    if b1_pld_zam_p != 0 and b2_pld_zam_p != 0:
        pow_wew_dolna.append([x_10, y_10])
        pow_wew_dolna.append([x_11, y_11])
    elif b1_pld_zam_p == 0 and b2_pld_zam_p != 0:
        pow_wew_dolna.append([x_10, y_10])
    elif b1_pld_zam_p != 0 and b2_pld_zam_p == 0:
        pow_wew_dolna.append([x_10, y_10])
        pow_wew_dolna.append([x_11, y_11])
    pow_wew_dolna.append([x_12, y_12])

    # Górna powierzchnia wewnętrzna:

    pow_wew_gorna = []

    # Punkt lewego zamocowania:
    if b1_plg_zam_l + b2_plg_zam_l != 0:
        robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_plg_zam_l] for i in pow_gorna]))
    else:
        robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_plg] for i in pow_gorna]))
    Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(robocza_1))
    robocza_2 = [x_7, y_7, x_7 - 2 * (x_3 - x_2), y_7 + 2 * (y_2 - y_3)]
    Robocza_2 = acad.model.AddLightWeightPolyline(aDouble(robocza_2))
    pkt_13 = find_intersections_2_selection([Robocza_1], [Robocza_2])
    Robocza_1.Delete()
    Robocza_2.Delete()
    x_13 = round(pkt_13[0][1][0], 12)
    y_13 = round(pkt_13[0][1][1], 12)
    pow_wew_gorna.append([x_13, y_13])

    # Punkt prawego zamocowania:
    if b1_plg_zam_p + b2_plg_zam_p != 0:
        robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_plg_zam_p] for i in pow_gorna]))
    else:
        robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_plg] for i in pow_gorna]))
    Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(robocza_1))
    robocza_2 = [x_12, y_12, x_12 + 2 * (x_5 - x_4), y_12 + 2 * (y_5 - y_4)]
    Robocza_2 = acad.model.AddLightWeightPolyline(aDouble(robocza_2))
    pkt_18 = find_intersections_2_selection([Robocza_1], [Robocza_2])
    Robocza_1.Delete()
    Robocza_2.Delete()
    x_18 = round(pkt_18[0][1][0], 12)
    y_18 = round(pkt_18[0][1][1], 12)

    # Punkty pośrednie:
    x_14 = round(x_13 + b1_plg_zam_l, 12)
    y_14 = y_13
    x_17 = round(x_18 - b1_plg_zam_p, 12)
    y_17 = y_18
    x_15 = round(x_14 + b2_plg_zam_l, 12)
    x_16 = round(x_17 - b2_plg_zam_p, 12)

    # Punkty płyty:
    robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_plg] for i in pow_gorna]))
    Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(robocza_1))
    robocza_2 = [x_15, y_g - 10, x_15, y_g + 10]
    Robocza_2 = acad.model.AddLightWeightPolyline(aDouble(robocza_2))
    robocza_3 = [x_16, y_g - 10, x_16, y_g + 10]
    Robocza_3 = acad.model.AddLightWeightPolyline(aDouble(robocza_3))
    pkt_15 = find_intersections_2_selection([Robocza_1], [Robocza_2])
    pkt_16 = find_intersections_2_selection([Robocza_1], [Robocza_3])
    Robocza_1.Delete()
    Robocza_2.Delete()
    Robocza_3.Delete()
    y_15 = round(pkt_15[0][1][1], 12)
    y_16 = round(pkt_16[0][1][1], 12)

    if b1_plg_zam_l != 0 and b2_plg_zam_l != 0:
        pow_wew_gorna.append([x_14, y_14])
        pow_wew_gorna.append([x_15, y_15])
    elif b1_plg_zam_l == 0 and b2_plg_zam_l != 0:
        pow_wew_gorna.append([x_15, y_15])
    elif b1_plg_zam_l != 0 and b2_plg_zam_l == 0:
        pow_wew_gorna.append([x_14, y_14])
        pow_wew_gorna.append([x_15, y_15])

    if b1_plg_zam_p != 0 and b2_plg_zam_p != 0:
        pow_wew_gorna.append([x_16, y_16])
        pow_wew_gorna.append([x_17, y_17])
    elif b1_plg_zam_p == 0 and b2_plg_zam_p != 0:
        pow_wew_gorna.append([x_16, y_16])
    elif b1_plg_zam_p != 0 and b2_plg_zam_p == 0:
        pow_wew_gorna.append([x_16, y_16])
        pow_wew_gorna.append([x_17, y_17])

    pow_wew_gorna.append([x_18, y_18])

    # Rysowanie konstrukcji:
    konstrukcja_lista = list(chain.from_iterable(pow_gorna + pow_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True
    konstrukcja_lista = list(chain.from_iterable(pow_wew_gorna + pow_wew_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True

    # ==================================================================================================================
    # OSIE I OPISY
    # ==================================================================================================================

    # Rzędna minimalna konstrukcji:
    x_os = round((x_3 + x_4) / 2, 8)
    y_os = y_3
    LWPline = acad.model.AddLightWeightPolyline(aDouble([x_os, y_os - 0.325, x_os, y_os + h_skrz + 0.2]))
    LWPline.LinetypeScale = 0.1
    LWPline.Layer = 'AII_M_osie główne'
    x_text = round(x_os - 0.075, 8)
    y_text = round(y_os + h_skrz / 2, 8)
    text = 'oś konstrukcji'
    Text = acad.model.AddText(text, aDouble(x_text, y_text, 0), 0.125)
    Text.Alignment = 1
    Text.TextAlignmentPoint = aDouble(x_text, y_text, 0)
    Text.Rotation = radians(90)
    Text.StyleName = 'AII_norm.'
    Text.Layer = 'AII_M_opis'
    Wym = acad.model.AddDimRotated(aDouble(x_os, y_os, 0),
                                   aDouble(x_os, y_os + h_skrz, 0),
                                   aDouble(x_os + 0.25, y_os, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    # ==================================================================================================================
    # WYMIARY DÓŁ
    # ==================================================================================================================

    # Wymiary poziome:

    # Poziomy wymiarów:
    y_min = y_os
    y_wym_1 = aDouble(0, y_min - 0.25, 0)
    y_wym_2 = aDouble(0, y_min - 0.5, 0)
    y_wym_3 = aDouble(0, y_min - 0.75, 0)
    y_wym_4 = aDouble(0, y_min - 1, 0)

    # Wymiary:
    wymiary_dol_1 = [[aDouble(x_3, y_3, 0), aDouble(x_os, y_os - 0.325, 0), 0],
                     [aDouble(x_os, y_os - 0.325, 0), aDouble(x_4, y_4, 0), 0]]
    wymiary_dol_2 = []
    wymiary_dol_3 = [[aDouble(x_1, y_1, 0), aDouble(x_os, y_os - 0.325, 0), 0],
                     [aDouble(x_os, y_os - 0.325, 0), aDouble(x_6, y_6, 0), 0]]
    wymiary_dol_4 = [[aDouble(x_1, y_1, 0), aDouble(x_6, y_6, 0), 0]]

    if x_2 == x_3:
        wymiary_dol_2.append([aDouble(x_1, y_1, 0), aDouble(x_3, y_3, 0), 0])
    else:
        wymiary_dol_2.append([aDouble(x_1, y_1, 0), aDouble(x_2, y_2, 0), 0])
        wymiary_dol_2.append([aDouble(x_2, y_2, 0), aDouble(x_3, y_3, 0), 0])

    wymiary_dol_2.append([aDouble(x_3, y_3, 0), aDouble(x_4, y_4, 0), 0])

    if x_4 == x_5:
        wymiary_dol_2.append([aDouble(x_4, y_4, 0), aDouble(x_6, y_6, 0), 0])
    else:
        wymiary_dol_2.append([aDouble(x_4, y_4, 0), aDouble(x_5, y_5, 0), 0])
        wymiary_dol_2.append([aDouble(x_5, y_5, 0), aDouble(x_6, y_6, 0), 0])

    # Wstawienie wymiarów:
    for index, wymiar in enumerate(wymiary_dol_1):
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_1, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        if index == 0:
            Wym.ExtLine2Suppress = True
        else:
            Wym.ExtLine1Suppress = True

    for wymiar in wymiary_dol_2:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_2, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for wymiar in wymiary_dol_3:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_3, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        Wym.ExtensionLineOffset = 7.0

    for wymiar in wymiary_dol_4:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_4, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # Wymiary poziome poszerzeń:

    # Płyta dolna - strona lewa:
    if b2_pld_zam_l != 0:
        if b1_pld_zam_l != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_7, y_7, 0),
                                           aDouble(x_8, y_8, 0),
                                           aDouble(x_7, y_7 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym = acad.model.AddDimRotated(aDouble(x_8, y_8, 0),
                                           aDouble(x_9, y_9, 0),
                                           aDouble(x_8, y_8 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_7, y_7, 0),
                                           aDouble(x_9, y_9, 0),
                                           aDouble(x_7, y_7 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
    else:
        if b1_pld_zam_l != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_7, y_7, 0),
                                           aDouble(x_8, y_8, 0),
                                           aDouble(x_7, y_7 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'

    # Płyta dolna - strona prawa:
    if b2_pld_zam_p != 0:
        if b1_pld_zam_p != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_11, y_11, 0),
                                           aDouble(x_12, y_12, 0),
                                           aDouble(x_11, y_11 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym = acad.model.AddDimRotated(aDouble(x_10, y_10, 0),
                                           aDouble(x_11, y_11, 0),
                                           aDouble(x_11, y_11 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_10, y_10, 0),
                                           aDouble(x_12, y_12, 0),
                                           aDouble(x_12, y_12 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
    else:
        if b1_pld_zam_p != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_11, y_11, 0),
                                           aDouble(x_12, y_12, 0),
                                           aDouble(x_11, y_11 + 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'

    # Płyta górna - strona lewa:
    if b2_plg_zam_l != 0:
        if b1_plg_zam_l != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_13, y_13, 0),
                                           aDouble(x_14, y_14, 0),
                                           aDouble(x_13, y_13 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5
            Wym = acad.model.AddDimRotated(aDouble(x_14, y_14, 0),
                                           aDouble(x_15, y_15, 0),
                                           aDouble(x_14, y_14 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_13, y_13, 0),
                                           aDouble(x_15, y_15, 0),
                                           aDouble(x_13, y_13 - 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
    else:
        if b1_plg_zam_l != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_13, y_13, 0),
                                           aDouble(x_14, y_14, 0),
                                           aDouble(x_13, y_13 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5

    # Płyta górna - strona prawa:
    if b2_plg_zam_p != 0:
        if b1_plg_zam_p != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_17, y_17, 0),
                                           aDouble(x_18, y_18, 0),
                                           aDouble(x_17, y_17 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5
            Wym = acad.model.AddDimRotated(aDouble(x_16, y_16, 0),
                                           aDouble(x_17, y_17, 0),
                                           aDouble(x_17, y_17 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5

        else:
            Wym = acad.model.AddDimRotated(aDouble(x_16, y_16, 0),
                                           aDouble(x_18, y_18, 0),
                                           aDouble(x_18, y_18 - 0.175, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
    else:
        if b1_plg_zam_p != 0:
            Wym = acad.model.AddDimRotated(aDouble(x_17, y_17, 0),
                                           aDouble(x_18, y_18, 0),
                                           aDouble(x_17, y_17 - 0.25, 0), 0)
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtensionLineOffset = 3.5

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

    # Płyta dolna:
    Wym = acad.model.AddDimRotated(aDouble(x_2, y_2, 0), aDouble(x_2, y_2 + h_wsp_zam, 0),
                                   aDouble(x_2 + 0.175, y_2, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(aDouble(x_5, y_5, 0), aDouble(x_5, y_5 + h_wsp_zam, 0),
                                   aDouble(x_5 - 0.175, y_5, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    if b1_pld_zam_l + b2_pld_zam_l != 0:
        if b1_pld_zam_l == 0:
            Wym = acad.model.AddDimRotated(aDouble(x_7, y_7, 0), aDouble(x_7, y_7 - h_pld_zam_l, 0),
                                           aDouble(x_7 + 0.175, y_7, 0), radians(90))
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_7, y_7, 0), aDouble(x_7, y_7 - h_pld_zam_l, 0),
                                           aDouble(x_7 + b1_pld_zam_l / 2, y_7, 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
    if b1_pld_zam_p + b2_pld_zam_p != 0:
        if b1_pld_zam_p == 0:
            Wym = acad.model.AddDimRotated(aDouble(x_12, y_12, 0), aDouble(x_12, y_12 - h_pld_zam_p, 0),
                                           aDouble(x_12 - 0.175, y_12, 0), radians(90))
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_12, y_12, 0), aDouble(x_12, y_12 - h_pld_zam_p, 0),
                                           aDouble(x_12 - b1_pld_zam_p / 2, y_12, 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(aDouble(x_os, y_os, 0), aDouble(x_os, y_os + h_pld, 0),
                                   aDouble(x_os, y_os, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym.ExtLine1Suppress = True
    Wym.ExtLine2Suppress = True

    # Płyta górna:
    if b1_plg_zam_l + b2_plg_zam_l != 0:
        Wym = acad.model.AddDimRotated(aDouble(x_13, y_13, 0), aDouble(x_13, y_13 + h_plg_zam_l, 0),
                                       aDouble(x_13 + 0.175, y_13, 0), radians(90))
    else:
        Wym = acad.model.AddDimRotated(aDouble(x_13, y_13, 0), aDouble(x_13, y_13 + h_plg, 0),
                                       aDouble(x_13 + 0.175, y_13, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    if b1_plg_zam_p + b2_plg_zam_p != 0:
        Wym = acad.model.AddDimRotated(aDouble(x_18, y_18, 0), aDouble(x_18, y_18 + h_plg_zam_p, 0),
                                       aDouble(x_18 - 0.175, y_18, 0), radians(90))
    else:
        Wym = acad.model.AddDimRotated(aDouble(x_18, y_18, 0), aDouble(x_18, y_18 + h_plg, 0),
                                       aDouble(x_18 - 0.175, y_18, 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    robocza_1 = list(chain.from_iterable(pow_wew_gorna))
    Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(robocza_1))
    Robocza_2 = acad.model.AddLightWeightPolyline(aDouble([x_os, y_os, x_os, y_os + 2 * h_skrz]))
    pkt = find_intersections_2_selection([Robocza_1], [Robocza_2])
    Robocza_1.Delete()
    Robocza_2.Delete()
    wym_h = round(y_os + h_skrz - pkt[0][1][1], 12)

    if wym_h == h_plg:
        Wym = acad.model.AddDimRotated(aDouble(x_os, pkt[0][1][1], 0), aDouble(x_os, y_os + h_skrz, 0),
                                       aDouble(x_os, pkt[0][1][1], 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        Wym.ExtLine1Suppress = True
        Wym.ExtLine2Suppress = True
    else:
        if b1_plg_zam_l == 0 and b2_plg_zam_l == 0:
            pass
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_15, y_15, 0), aDouble(x_15, y_15 + h_plg, 0),
                                           aDouble(x_15 - 0.175, y_15, 0), radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
        if b1_plg_zam_p == 0 and b2_plg_zam_p == 0:
            pass
        else:
            Wym = acad.model.AddDimRotated(aDouble(x_16, y_16, 0), aDouble(x_16, y_16 + h_plg, 0),
                                           aDouble(x_16 + 0.175, y_16, 0), radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'

    # Środniki:
    x1 = round((x_7 + x_13) / 2, 12)
    y1 = round((y_7 + y_13) / 2, 12)
    if x_2 == x_3:
        angle = radians(90)
    else:
        angle = atan((y_2 - y_3) / (x_3 - x_2))
    x2 = round(x1 - t_sr * cos(radians(90) - angle), 12)
    y2 = round(y1 - t_sr * sin(radians(90) - angle), 12)
    Wym = acad.model.AddDimRotated(aDouble(x2, y2, 0), aDouble(x1, y1, 0),
                                   aDouble(x2, y2, 0), radians(90) - angle)
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym.ExtLine1Suppress = True
    Wym.ExtLine2Suppress = True

    # Środniki:
    x1 = round((x_12 + x_18) / 2, 12)
    y1 = round((y_12 + y_18) / 2, 12)
    if x_4 == x_5:
        angle = radians(90)
    else:
        angle = atan((y_5 - y_4) / (x_5 - x_4))
    x2 = round(x1 + t_sr * cos(radians(90) - angle), 12)
    y2 = round(y1 - t_sr * sin(radians(90) - angle), 12)
    Wym = acad.model.AddDimRotated(aDouble(x1, y1, 0), aDouble(x2, y2, 0),
                                   aDouble(x1, y1, 0), angle - radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym.ExtLine1Suppress = True
    Wym.ExtLine2Suppress = True

    print(f'Obiekt {obiekt} - ustrój betonowy, skrzynkowy narysowany!')
