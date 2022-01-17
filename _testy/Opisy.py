from pyautocad import aDouble, Autocad
import pandas as pd
from math import isnan, atan, fabs, radians, sin, cos
from itertools import chain
import win32com.client
from functools import wraps
from time import time
import Pobieranie_danych
from Funkcje_podstawowe import *

acad = Autocad()

version = acad.doc.GetVariable("ACADVER")
color = acad.app.GetInterfaceObject(f'AutoCAD.ACCmColor.{version[0:2]}')
acad_32 = win32com.client.Dispatch("AutoCAD.Application")


@speed_test
def opis(opis_gora):

    # Pobranie danych:
    obiekt=Pobieranie_danych.obiekt
    delta_y=Pobieranie_danych.delta_y

    # ==================================================================================================================
    # WYMIARY GÓRA
    # ==================================================================================================================

    if len(opis_gora) == 1:
        y_wym = opis_gora[0][0]
        wymiary_gora_1 = opis_gora[0][1]
        wymiary_gora_2 = opis_gora[0][2]
        wymiary_gora_3 = opis_gora[0][3]
        wymiary_gora_4 = opis_gora[0][4]
    else:
        y_wym = max(opis_gora[0][0], opis_gora[1][0])
        wymiary_gora_1 = opis_gora[0][1] + opis_gora[1][1]
        wymiary_gora_2 = opis_gora[0][2] + opis_gora[1][2]
        wymiary_gora_3 = opis_gora[0][3] + opis_gora[1][3]
        wymiary_gora_4 = opis_gora[0][4] + opis_gora[1][4]

    y_wym_1 = aDouble(0, y_wym, 0)
    y_wym_2 = aDouble(0, y_wym + 0.25, 0)
    y_wym_3 = aDouble(0, y_wym + 0.5, 0)
    y_wym_4 = aDouble(0, y_wym + 0.75, 0)
    y_wym_5 = aDouble(0, y_wym + 1, 0)

    for wymiar in wymiary_gora_1:
        if wymiar[0][1] == 1000:
            wymiar[0][1] = y_wym - 0.175
        if wymiar[1][1] == 1000:
            wymiar[1][1] = y_wym - 0.175

    for wymiar in wymiary_gora_1:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_1, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'
        if wymiar[2] == 0:
            pass
        else:
            Wym.TextOverride = f'<>\X{wymiar[2]}'

    for wymiar in wymiary_gora_2:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_2, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for wymiar in wymiary_gora_3:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_3, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    for wymiar in wymiary_gora_4:
        Wym = acad.model.AddDimRotated(wymiar[0], wymiar[1], y_wym_4, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    if len(opis_gora) == 2:
        if len(wymiary_gora_4) == 0:
            y_wym_dol = y_wym_3
            y_wym_gora = y_wym_4
        else:
            y_wym_dol = y_wym_4
            y_wym_gora = y_wym_5

        Wym = acad.model.AddDimRotated(aDouble(opis_gora[0][7], opis_gora[0][8], 0),
                                       aDouble(opis_gora[1][5], opis_gora[1][6], 0), y_wym_dol, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

        Wym = acad.model.AddDimRotated(aDouble(opis_gora[0][5], opis_gora[0][6], 0),
                                       aDouble(opis_gora[1][7], opis_gora[1][8], 0), y_wym_gora, 0)
        Wym.Layer = 'AII_M_wymiary'
        Wym.StyleName = 'A2_M50MM'

    # ==================================================================================================================
    # OŚ NIWELETY
    # ==================================================================================================================
    osie = []
    if len(opis_gora) == 1:
        osie.append(opis_gora[0][11] + [y_wym - 0.175])
    else:
        osie.append(opis_gora[0][11] + [y_wym - 0.175])
        osie.append(opis_gora[1][11] + [y_wym - 0.175])
    for os in osie:
        LWPline = acad.model.AddLightWeightPolyline(aDouble(os))
        LWPline.LinetypeScale = 0.1
        LWPline.Layer = 'AII_M_osie główne'

    # ==================================================================================================================
    # TYTUŁ
    # ==================================================================================================================
    if len(opis_gora) == 2:
        x_tyt = round((opis_gora[0][5] + opis_gora[1][7]) / 2, 8)
    else:
        x_tyt = round((opis_gora[0][5] + opis_gora[0][7]) / 2, 8)

    if len(wymiary_gora_4) == 0:
        y_tyt_1 = y_wym + 1.35
    else:
        y_tyt_1 = y_wym + 1.6

    y_tyt_2 = y_tyt_1 + 0.4
    ins_point_1 = win32_point(x_tyt, y_tyt_1, 0)
    ins_point_2 = win32_point(x_tyt, y_tyt_2, 0)
    Block = acad_32.ActiveDocument.ModelSpace.InsertBlock(ins_point_1, 'Tytuł', 1, 1, 1, 0)
    Block.Layer = 'AII_M_opis'
    Block = acad_32.ActiveDocument.ModelSpace.InsertBlock(ins_point_2, 'Obiekt', 1, 1, 1, 0)
    atr = Block.GetAttributes()
    atr[0].TextString = f'{obiekt}'
    Block.Layer = 'AII_M_pomoc'

    # ==================================================================================================================
    # KOTY
    # ==================================================================================================================

    if len(opis_gora) == 1:
        koty = opis_gora[0][9]
    else:
        koty_lewe = opis_gora[0][9]
        koty_prawe = opis_gora[1][9]
        if delta_y != 0:
            for kota in koty_lewe:
                kota[4] = f'{kota[4]} NL'
            for kota in koty_prawe:
                kota[4] = f'{kota[4]} NP'
        koty = koty_lewe + koty_prawe

    for kota in koty:
        ins_point = win32_point(kota[0], kota[1], kota[2])
        Block = acad_32.ActiveDocument.ModelSpace.InsertBlock(ins_point, '00kota3', 0.05, 0.05, 0.05, 0)
        atr = Block.GetAttributes()
        atr[0].TextString = kota[4]

        if kota[4] == '%%p0.000 NL':
            x = round(kota[0] + 0.02389966, 8)
            y = round(kota[1] + 0.28462685, 8)
            if delta_y > 0:
                tekst = f'{-delta_y:.3f} NP'
            else:
                tekst = f'+{-delta_y:.3f} NP'
            Text = acad.model.AddText(tekst, aDouble(x, y, 0), 0.1)
            Text.StyleName = 'AII_norm.'
            Text.Layer = 'AII_M_opis'
            Text.ScaleFactor = 0.85

        if kota[4] == '%%p0.000 NP':
            x = round(kota[0] + 0.02389966, 8)
            y = round(kota[1] + 0.28462685, 8)
            if delta_y > 0:
                tekst = f'+{delta_y:.3f} NL'
            else:
                tekst = f'-{-delta_y:.3f} NL'
            Text = acad.model.AddText(tekst, aDouble(x, y, 0), 0.1)
            Text.StyleName = 'AII_norm.'
            Text.Layer = 'AII_M_opis'
            Text.ScaleFactor = 0.85

        if kota[3] == -1:
            Block2 = Block.Mirror(win32_point(kota[0], kota[1], 0), win32_point(kota[0], kota[1] + 1, 0))
            Block.Delete()
            Block = Block2
            if kota[4] in ['%%p0.000 NL', '%%p0.000 NP']:
                Text2 = Text.Mirror(win32_point(kota[0], kota[1], 0), win32_point(kota[0], kota[1] + 1, 0))
                Text.Delete()
                Text = Text2

        Block.Layer = 'AII_M_koty wysokościowe'

        if kota[4] in ['%%p0.000 NL', '%%p0.000 NP']:
            Text.StyleName = 'AII_norm.'
            Text.Layer = 'AII_M_opis'
            Text.ScaleFactor = 0.85

    print(f'Obiekt {obiekt} - opis gotowy!')
