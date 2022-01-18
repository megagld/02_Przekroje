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
def rysowanie_konstrukcja_belkowy(pow_gorna):
    # Pobranie danych:
    obiekt=Pobieranie_danych.obiekt
    konstrukcja = Pobieranie_danych.konstrukcja    
    
    h_wsp_kr = konstrukcja['B_WSP - h'][0]
    h_wsp_zam = konstrukcja['B_WSP - h zam'][0]
    h_pl = konstrukcja['B_PL - h'][0]
    h_pl_zam = konstrukcja['B_PL - h zam'][0]
    szer_zam = konstrukcja['B_PL - szer zam'][0]
    h_dzw = konstrukcja['B_DZW - h'][0]
    b_dzw = konstrukcja['B_DZW - b'][0]
    n_dzw = konstrukcja['B_DZW - n'][0]
    roz_dzw = konstrukcja['B_DZW - roz'][0]
    skos_dzw = konstrukcja['B_DZW - skos'][0]

    # Dolna powierzchnia konstrukcji:
    x_kon = round((pow_gorna[-1][0] + pow_gorna[0][0]) / 2, 8)

    x_osie = []
    x_pl_zam = []
    x_pl_pl = []
    x_wsp_zam = []
    x_wsp_kr = [pow_gorna[0][0], pow_gorna[-1][0]]
    xy_dzw = []

    # Obliczenie położenia osi dźwigarów:
    if n_dzw % 2 == 0:
        for i in range(int(n_dzw / 2)):
            if i == 0:
                x_p = x_kon + roz_dzw / 2
                x_l = x_kon - roz_dzw / 2
                x_osie.append(round(x_l, 6))
                x_osie.append(round(x_p, 6))
            else:
                x_p += roz_dzw
                x_l -= roz_dzw
                x_osie.append(round(x_l, 6))
                x_osie.append(round(x_p, 6))
    else:
        for i in range(int((n_dzw + 1) / 2)):
            if i == 0:
                x_p = x_kon
                x_l = x_kon
                x_osie.append(round(x_kon, 6))
            else:
                x_p += roz_dzw
                x_l -= roz_dzw
                x_osie.append(round(x_l, 6))
                x_osie.append(round(x_p, 6))

    x_osie = sorted(x_osie)
    xy_osie = []

    # Obliczenie punktów charakterystycznych:
    for i, x_os in enumerate(x_osie):
        # Współrzędna y dźwigara:
        for j in range(len(pow_gorna)):
            x2 = pow_gorna[j][0]
            if x_os <= x2:
                index = j
                break
        x1 = pow_gorna[index - 1][0]
        y1 = pow_gorna[index - 1][1]
        x2 = pow_gorna[index][0]
        y2 = pow_gorna[index][1]
        y_dzw = round(y1 + (y2 - y1) * (x_os - x1) / (x2 - x1), 6)
        xy_osie.append([x_os, y_dzw])

        # Współrzędne x punktów charakterystycznych:
        x_dzw_l = round(x_os - b_dzw / 2, 6)
        x_dzw_zam_l = round(x_dzw_l - skos_dzw, 6)
        x_dzw_pl_l = round(x_dzw_zam_l - szer_zam, 6)
        x_dzw_p = round(x_os + b_dzw / 2, 6)
        x_dzw_zam_p = round(x_dzw_p + skos_dzw, 6)
        x_dzw_pl_p = round(x_dzw_zam_p + szer_zam, 6)

        # Dopisanie współrzędnych do odpowiednich list:
        xy_dzw.append([x_dzw_l, y_dzw])
        xy_dzw.append([x_dzw_p, y_dzw])
        if i == 0:
            x_wsp_zam.append(x_dzw_zam_l)
            x_pl_zam.append(x_dzw_zam_p)
            if h_pl_zam != h_pl:
                x_pl_pl.append(x_dzw_pl_p)
        elif i == len(x_osie) - 1:
            x_pl_zam.append(x_dzw_zam_l)
            x_wsp_zam.append(x_dzw_zam_p)
            if h_pl_zam != h_pl:
                x_pl_pl.append(x_dzw_pl_l)
        else:
            x_pl_zam.append(x_dzw_zam_l)
            x_pl_zam.append(x_dzw_zam_p)
            if h_pl_zam != h_pl:
                x_pl_pl.append(x_dzw_pl_l)
                x_pl_pl.append(x_dzw_pl_p)

    # Obliczenie współrzędnych y dla punktów charakterystycznych:
    pow_dolna_x = x_pl_zam + x_pl_pl + x_wsp_zam + x_wsp_kr
    pow_dolna = [i for i in xy_dzw]

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

    for pkt in pow_dolna:
        if pkt in xy_dzw:
            y = round(pkt[1] - h_dzw, 6)
        elif pkt[0] in x_wsp_kr:
            y = round(pkt[1] - h_wsp_kr, 6)
        elif pkt[0] in x_wsp_zam:
            y = round(pkt[1] - h_wsp_zam, 6)
        elif pkt[0] in x_pl_pl:
            y = round(pkt[1] - h_pl, 6)
        elif pkt[0] in x_pl_zam:
            y = round(pkt[1] - h_pl_zam, 6)
        pkt[1] = y

    konstrukcja_lista = list(chain.from_iterable(pow_gorna + pow_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True

    # ==================================================================================================================
    # OSIE I OPISY
    # ==================================================================================================================

    # Rzędna minimalna konstrukcji:
    y_min = min([i[1] for i in pow_dolna])

    # Wstawianie osi:
    for index, pkt in enumerate(xy_osie):
        LWPline = acad.model.AddLightWeightPolyline(aDouble([pkt[0], y_min - 0.575, pkt[0], pkt[1] + 0.2]))
        LWPline.LinetypeScale = 0.1
        LWPline.Layer = 'AII_M_osie główne'
        x_text = pkt[0] - 0.075
        y_text = pkt[1] - h_dzw / 2
        text = f'oś dźwigara {chr(97 + index).upper()}'
        Text = acad.model.AddText(text, aDouble(x_text, y_text, 0), 0.125)
        Text.Alignment = 1
        Text.TextAlignmentPoint = aDouble(x_text, y_text, 0)
        Text.Rotation = radians(90)
        Text.StyleName = 'AII_norm.'
        Text.Layer = 'AII_M_opis'
        Wym = acad.model.AddDimRotated(aDouble(pkt[0], pkt[1] - h_dzw, 0),
                                       aDouble(pkt[0], pkt[1], 0),
                                       aDouble(pkt[0] + 0.25, pkt[1], 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # ==================================================================================================================
    # WYMIARY DÓŁ
    # ==================================================================================================================

    # Wymiary poziome:

    # Poziomy wymiarów:
    y_wym_1 = aDouble(0, y_min - 0.25, 0)
    y_wym_2 = aDouble(0, y_min - 0.5, 0)
    y_wym_3 = aDouble(0, y_min - 0.75, 0)

    # Wymiary:
    wymiary_dol_1 = []
    wymiary_dol_2 = []
    for i in range(len(pow_dolna) - 1):
        x1 = pow_dolna[i][0]
        y1 = pow_dolna[i][1]
        x2 = pow_dolna[i + 1][0]
        y2 = pow_dolna[i + 1][1]

        # Zmiana współrzędnej zamocowania, jeśli powierzchnie boczne dźwigarów są pionowe:
        if x1 in [i[0] for i in xy_dzw]:
            y1 = [i[1] for i in xy_dzw if i[0] == x1][0]
        if x2 in [i[0] for i in xy_dzw]:
            y2 = [i[1] for i in xy_dzw if i[0] == x2][0]
        # Warunek na wymiar skosu (brak jeśli powierzchnie boczne dźwigarów są pionowe):
        if x1 != x2:
            wymiary_dol_1.append([aDouble(x1, y1, 0), aDouble(x2, y2, 0), 0])

    for i in range(len(x_osie) - 1):
        x1 = x_osie[i]
        x2 = x_osie[i + 1]
        y = y_min - 0.575
        wymiary_dol_2.append([aDouble(x1, y, 0), aDouble(x2, y, 0), 0])

    wymiary_dol_2.insert(0, [aDouble(pow_dolna[0][0], pow_dolna[0][1], 0), aDouble(x_osie[0], y_min - 0.575, 0), 0])
    wymiary_dol_2.append([aDouble(x_osie[-1], y_min - 0.575, 0), aDouble(pow_dolna[-1][0], pow_dolna[-1][1], 0), 0])
    x_lewy_skr_dol = pow_dolna[0][0]
    y_lewy_skr_dol = pow_dolna[0][1]
    x_prawy_skr_dol = pow_dolna[-1][0]
    y_prawy_skr_dol = pow_dolna[-1][1]
    wymiary_dol_3 = [[aDouble(x_lewy_skr_dol, y_lewy_skr_dol, 0), aDouble(x_prawy_skr_dol, y_prawy_skr_dol, 0), 0]]

    # Wstawienie wymiarów:
    for wymiar in wymiary_dol_1:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_1, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for index, wymiar in enumerate(wymiary_dol_2):
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_2, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        if index == 0:
            Wym.ExtLine2Suppress = True
        elif index == len(wymiary_dol_2) - 1:
            Wym.ExtLine1Suppress = True
        else:
            Wym.ExtLine1Suppress = True
            Wym.ExtLine2Suppress = True

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
    x_zam = sorted(x_wsp_zam + x_pl_zam)
    xy_zam = []
    for x in x_zam:
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
        xy_zam.append([round(x, 6), round(y, 6)])
    for index, pkt in enumerate(xy_zam):
        x = pkt[0]
        if x in x_pl_zam:
            y1 = pkt[1] - h_pl_zam
        else:
            y1 = pkt[1] - h_wsp_zam
        y2 = pkt[1]
        if index % 2 == 0:
            x_wym = x + 0.175
        else:
            x_wym = x - 0.175
        Wym = acad.model.AddDimRotated(aDouble(x, y1, 0), aDouble(x, y2, 0), aDouble(x_wym, y1, 0), radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # Płyty:
    x_pl = sorted(x_pl_pl)
    xy_pl = []
    for x in x_pl:
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
        xy_pl.append([round(x, 6), round(y, 6)])

    for i in range(int(len(xy_pl) / 2)):
        x1 = xy_pl[2 * i][0]
        y1 = xy_pl[2 * i][1]
        x2 = xy_pl[2 * i + 1][0]
        y2 = xy_pl[2 * i + 1][1]
        Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
        Robocza_2 = acad.model.AddLightWeightPolyline(aDouble([(x1 + x2) / 2, y1 - 10, (x1 + x2) / 2, y1 + 10]))
        pkty = find_intersections_2_selection([Robocza_1], [Robocza_2])
        Robocza_1.Delete()
        Robocza_2.Delete()
        y3 = pkty[0][1][1]
        y4 = pkty[0][1][4]
        x3 = (x1 + x2) / 2
        if round(y3 - y4, 6) == h_pl:
            Wym = acad.model.AddDimRotated(aDouble(x3, y4, 0), aDouble(x3, y3, 0), aDouble(x3, y3, 0), radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtLine1Suppress = True
            Wym.ExtLine2Suppress = True
        else:
            Wym = acad.model.AddDimRotated(aDouble(x1, y1 - h_pl, 0), aDouble(x1, y1, 0), aDouble(x1, y1, 0),
                                           radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtLine1Suppress = True
            Wym.ExtLine2Suppress = True
            Wym = acad.model.AddDimRotated(aDouble(x2, y2 - h_pl, 0), aDouble(x2, y2, 0), aDouble(x2, y2, 0),
                                           radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
            Wym.ExtLine1Suppress = True
            Wym.ExtLine2Suppress = True

    print(f'Obiekt {obiekt} - ustrój betonowy, belkowy narysowany!')

