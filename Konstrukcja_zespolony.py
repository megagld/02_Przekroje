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
def rysowanie_konstrukcja_zespolony(pow_gorna):
    # konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56], sheet_name=sheet)
 
    obiekt=Pobieranie_danych.obiekt
    konstrukcja = Pobieranie_danych.konstrukcja   

    h_wsp = konstrukcja['Z_WSP - h'][0]
    h_pl = konstrukcja['Z_PL - h'][0]
    skos_pl = konstrukcja['Z_PL - skos'][0]
    skos_h = konstrukcja['Z_PL - skos h'][0]
    h_dzw = konstrukcja['Z_DZW - h'][0]
    n_dzw = konstrukcja['Z_DZW - n'][0]
    roz_dzw = konstrukcja['Z_DZW - roz'][0]
    b_pg = round(konstrukcja['Z_PG - b'][0] / 1000, 8)
    t_pg = round(konstrukcja['Z_PG - t'][0] / 1000, 8)
    b_pd = round(konstrukcja['Z_PD - b'][0] / 1000, 8)
    t_pd = round(konstrukcja['Z_PD - t'][0] / 1000, 8)
    t_sr = round(konstrukcja['Z_ŚR - t'][0] / 1000, 8)
    z_od_g = round(konstrukcja['Z_Ż - od g'][0] / 1000, 8)
    z_od_d = round(konstrukcja['Z_Ż - od d'][0] / 1000, 8)

    # ==================================================================================================================
    # PŁYTA ŻELBETOWA
    # ==================================================================================================================

    # Dolna powierzchnia konstrukcji:
    x_kon = round((pow_gorna[-1][0] + pow_gorna[0][0]) / 2, 8)
    x_osie = []
    x_pl = []
    pow_dolna = [[pow_gorna[0][0], round(pow_gorna[0][1] - h_wsp, 8)],
                 [pow_gorna[-1][0], round(pow_gorna[-1][1] - h_wsp, 8)]]

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
    xy_dzw = []

    for x in x_osie:
        if skos_pl == 'N':
            x_l = round(x - b_pg / 2, 8)
            x_p = round(x + b_pg / 2, 8)
        else:
            x_l = round(x - b_pg / 2 - skos_h, 8)
            x_p = round(x + b_pg / 2 + skos_h, 8)
            x_dzw_l = round(x - b_pg / 2, 8)
            x_dzw_p = round(x + b_pg / 2, 8)

        for i in range(len(pow_gorna)):
            x2 = pow_gorna[i][0]
            if x_l <= x2:
                index = i
                break
        x1 = pow_gorna[index - 1][0]
        y1 = pow_gorna[index - 1][1]
        x2 = pow_gorna[index][0]
        y2 = pow_gorna[index][1]
        y_l = round(y1 + (y2 - y1) * (x_l - x1) / (x2 - x1) - h_pl, 6)

        for i in range(len(pow_gorna)):
            x2 = pow_gorna[i][0]
            if x_p <= x2:
                index = i
                break
        x1 = pow_gorna[index - 1][0]
        y1 = pow_gorna[index - 1][1]
        x2 = pow_gorna[index][0]
        y2 = pow_gorna[index][1]
        y_p = round(y1 + (y2 - y1) * (x_p - x1) / (x2 - x1) - h_pl, 6)

        y_min = min(y_l, y_p)
        if skos_pl == 'N':
            y_dzw = round((y_l + y_p) / 2, 8)
            y_l = y_dzw
            y_p = y_dzw
        else:
            y_dzw = round(y_min - skos_h, 8)
            pow_dolna.append([x_dzw_l, y_dzw])
            pow_dolna.append([x_dzw_p, y_dzw])
            robocza_1 = list(chain.from_iterable([[i[0], i[1] - h_pl] for i in pow_gorna]))
            Robocza_1 = acad.model.AddLightWeightPolyline(aDouble(robocza_1))
            if y_l <= y_p:
                Robocza_2 = acad.model.AddLightWeightPolyline(aDouble([x_dzw_p, y_dzw, x_dzw_p + 1, y_dzw + 1]))
                pkt = find_intersections_2_selection([Robocza_1], [Robocza_2])
                x_p = pkt[0][1][0]
                y_p = pkt[0][1][1]
            else:
                Robocza_2 = acad.model.AddLightWeightPolyline(aDouble([x_dzw_l - 1, y_dzw + 1, x_dzw_l, y_dzw]))
                pkt = find_intersections_2_selection([Robocza_1], [Robocza_2])
                x_l = pkt[0][1][0]
                y_l = pkt[0][1][1]

            Robocza_1.Delete()
            Robocza_2.Delete()

        pow_dolna.append([x_l, y_l])
        pow_dolna.append([x_p, y_p])
        xy_dzw.append([x, y_dzw])

    pow_dolna = sorted(pow_dolna)

    if skos_pl == 'T':
        del pow_dolna[1]
        del pow_dolna[-2]
        for i in range(int(n_dzw - 1)):
            x1 = pow_dolna[4 * i + 3][0]
            x2 = pow_dolna[4 * i + 4][0]
            for pkt in pow_gorna:
                if x1 < pkt[0] < x2:
                    pow_dolna.append([pkt[0], round(pkt[1] - h_pl, 8)])
    else:
        for i in range(int(n_dzw - 1)):
            x1 = pow_dolna[2 * i + 2][0]
            x2 = pow_dolna[2 * i + 3][0]
            for pkt in pow_gorna:
                if x1 < pkt[0] < x2:
                    pow_dolna.append([pkt[0], round(pkt[1] - h_pl, 8)])
    pow_dolna = sorted(pow_dolna)

    konstrukcja_lista = list(chain.from_iterable(pow_gorna + pow_dolna[::-1]))
    LWPline = acad.model.AddLightWeightPolyline(aDouble(konstrukcja_lista))
    LWPline.Layer = 'AII_M_konstrukcja beton'
    LWPline.Closed = True

    y_min = min([i[1] for i in xy_dzw]) - h_dzw
    # ==================================================================================================================
    # DŹWIGARY STALOWE
    # ==================================================================================================================
    for index, pkt in enumerate(xy_dzw):
        # Punkty charakterystyczne:
        x_pg_l = round(pkt[0] - b_pg / 2, 8)
        x_pg_p = round(pkt[0] + b_pg / 2, 8)
        x_pd_l = round(pkt[0] - b_pd / 2, 8)
        x_pd_p = round(pkt[0] + b_pd / 2, 8)
        x_sr_l = round(pkt[0] - t_sr / 2, 8)
        x_sr_p = round(pkt[0] + t_sr / 2, 8)
        y_pg_g = round(pkt[1], 8)
        y_pg_d = round(pkt[1] - t_pg, 8)
        y_pd_g = round(pkt[1] - h_dzw + t_pd, 8)
        y_pd_d = round(pkt[1] - h_dzw, 8)

        # Rysowanie dźwigarów:
        dzwigar_lista = [x_pg_p, y_pg_g, x_pg_p, y_pg_d, x_sr_p, y_pg_d, x_sr_p, y_pd_g, x_pd_p, y_pd_g, x_pd_p, y_pd_d,
                         x_pd_l, y_pd_d, x_pd_l, y_pd_g, x_sr_l, y_pd_g, x_sr_l, y_pg_d, x_pg_l, y_pg_d, x_pg_l, y_pg_g]
        LWPline = acad.model.AddLightWeightPolyline(aDouble(dzwigar_lista))
        LWPline.Layer = 'AII_M_konstrukcja stal'
        LWPline.Closed = True

        # Rysowanie żeber:
        LWPline = acad.model.AddLightWeightPolyline(aDouble([round(x_pg_l + z_od_g, 8), y_pg_d,
                                                            round(x_pd_l + z_od_d, 8), y_pd_g]))
        LWPline.Layer = 'AII_M_konstrukcja stal'
        color.ColorIndex = 1
        LWPline.TrueColor = color
        LWPline = acad.model.AddLightWeightPolyline(aDouble([round(x_pg_p - z_od_g, 8), y_pg_d,
                                                            round(x_pd_p - z_od_d, 8), y_pd_g]))
        LWPline.Layer = 'AII_M_konstrukcja stal'
        color.ColorIndex = 1
        LWPline.TrueColor = color

        # Rysowanie osi:
        LWPline = acad.model.AddLightWeightPolyline(aDouble([pkt[0], y_min - 0.575, pkt[0], pkt[1] + 0.1]))
        LWPline.LinetypeScale = 0.1
        LWPline.Layer = 'AII_M_osie główne'

        # Wstawianie wymiarów:
        pkt_1 = aDouble(x_pd_p, y_pd_d, 0)
        pkt_2 = aDouble(x_pg_p, y_pg_g, 0)
        pkt_3 = aDouble(max(x_pd_p, x_pg_p) + 0.175, y_pd_d, 0)
        Wym = acad.model.AddDimRotated(pkt_1, pkt_2, pkt_3, radians(90))
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        if skos_pl == 'N':
            pkt_1 = aDouble(x_pg_p, y_pg_g, 0)
            pkt_2 = aDouble(pkt[0], pkt[1] + h_pl, 0)
            Wym = acad.model.AddDimRotated(pkt_1, pkt_2, pkt_3, radians(90))
            Wym.Layer = 'AII_M_wymiary'
            Wym.StyleName = 'A2_M50MM'
        else:
            if index > 0:
                x = round(pkt[0] - roz_dzw / 2, 8)
                y1 = round(pkt[1] - 20, 8)
                y2 = round(pkt[1] + 20, 8)
                Robocza_1 = acad.model.AddLightWeightPolyline(aDouble([x, y1, x, y2]))
                robocza_2 = list(chain.from_iterable(pow_dolna))
                Robocza_2 = acad.model.AddLightWeightPolyline(aDouble(robocza_2))
                pkt_1 = find_intersections_2_selection([Robocza_1], [Robocza_2])
                pkt_2 = aDouble(pkt_1[0][1][0], pkt_1[0][1][1] + h_pl, 0)
                pkt_1 = aDouble(pkt_1[0][1][0], pkt_1[0][1][1], 0)
                Robocza_1.Delete()
                Robocza_2.Delete()
                Wym = acad.model.AddDimRotated(pkt_1, pkt_2, pkt_2, radians(90))
                Wym.Layer = 'AII_M_wymiary'
                Wym.StyleName = 'A2_M50MM'
                Wym.ExtLine1Suppress = True
                Wym.ExtLine2Suppress = True

    # ==================================================================================================================
    # WYMIARY DÓŁ
    # ==================================================================================================================
    y_wym_1 = aDouble(0, y_min - 0.25, 0)
    y_wym_2 = aDouble(0, y_min - 0.5, 0)
    y_wym_3 = aDouble(0, y_min - 0.75, 0)

    wymiary_dol_1 = []
    wymiary_dol_2 = []

    for index, pkt in enumerate(xy_dzw):
        x1 = round(pkt[0] - b_pd / 2, 8)
        x2 = round(pkt[0] + b_pd / 2, 8)
        y = round(pkt[1] - h_dzw, 8)
        wymiary_dol_1.append([aDouble(x1, y, 0), aDouble(x2, y, 0), 0])
        if index > 0:
            x3 = xy_dzw[index - 1][0]
            x4 = xy_dzw[index][0]
            y = round(y_min - 0.575, 8)
            wymiary_dol_2.append([aDouble(x3, y, 0), aDouble(x4, y, 0), 0])

    wymiary_dol_2.insert(0, [aDouble(pow_dolna[0][0], pow_dolna[0][1], 0), aDouble(xy_dzw[0][0], xy_dzw[0][1], 0), 0])
    wymiary_dol_2.append([aDouble(xy_dzw[-1][0], xy_dzw[-1][1], 0), aDouble(pow_dolna[-1][0], pow_dolna[-1][1], 0), 0])

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
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[0][0], pow_gorna[0][1] - h_wsp, 0),
                                   aDouble(pow_gorna[0][0], pow_gorna[0][1], 0),
                                   aDouble(pow_gorna[0][0] - 0.175, pow_gorna[0][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'
    Wym = acad.model.AddDimRotated(aDouble(pow_gorna[-1][0], pow_gorna[-1][1] - h_wsp, 0),
                                   aDouble(pow_gorna[-1][0], pow_gorna[-1][1], 0),
                                   aDouble(pow_gorna[-1][0] + 0.3, pow_gorna[-1][1], 0), radians(90))
    Wym.Layer = 'AII_M_wymiary'
    Wym.StyleName = 'A2_M50MM'

    print(f'Obiekt {obiekt} - ustrój zespolony narysowany!')

