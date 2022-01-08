from pyautocad import aDouble, Autocad
import pandas as pd
from math import isnan, atan, fabs, radians, sin, cos
from itertools import chain
import win32com.client
import pythoncom
from functools import wraps
from time import time

acad = Autocad()
version = acad.doc.GetVariable("ACADVER")
color = acad.app.GetInterfaceObject(f'AutoCAD.ACCmColor.{version[0:2]}')
acad_32 = win32com.client.Dispatch("AutoCAD.Application")


def speed_test(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        start_time = time()
        result = fn(*args, **kwargs)
        end_time = time()
        delta_time = round(end_time - start_time, 3)
        if delta_time > 60:
            minutes, seconds = delta_time // 60, delta_time % 60
            print(f"Czas: {int(minutes)}min {round(seconds, 0)}s")
        else:
            print(f"Czas: {round(delta_time, 2)}s")
        return result

    return wrapper


def find_intersections_2_selection(selection_set1, selection_set2):
    points_list = []
    for obj in selection_set1:
        for next_obj in selection_set2:
            intersection_point = obj.IntersectWith(next_obj, 0)
            points_list.append([0, intersection_point])
    return points_list


def win32_point(x, y, z):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


@speed_test
def rysowanie_przekroj_ruchowy(x_g, y_g, file, sheet):
    # Pobranie danych:
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
    konstrukcja = pd.read_excel(file, usecols=[43], sheet_name=sheet)

    # ==================================================================================================================
    # JEZDNIA
    # ==================================================================================================================

    # Punkt 0,0:
    jezdnia = [[x_g, y_g]]

    # Punkty:
    for index, row in pasy_lewe.iterrows():
        if index == 0:
            x = x_g
            y = y_g
        if isnan(pasy_lewe.iloc[index, 0]) or pasy_lewe.iloc[index, 0] == 0:
            break
        x -= row['PL - szer']
        y += row['PL - szer'] * row['PL - spadek'] / 100
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

    # Punkt początkowy:
    x_ch_lewy_1 = x_kraw_l - 0.2
    y_ch_lewy_1 = y_kraw_l + 0.14

    # Określenie szerokości:
    if bar_lewa_rodz == 'linowa':
        bar_lewa_szer = 0.15
    else:
        bar_lewa_szer = 0.40

    if bal_lewa_rodz == 'B':
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
        if bar_lewa_rodz == 'linowa':
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
            name = f'Barieroporęcz_{bar_lewa_rodz}'
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
            name = f'Barieroporęcz_{bar_lewa_rodz}'
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
        if bal_lewa_rodz == 'B':
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

    # Punkt początkowy:
    x_ch_prawy_1 = x_kraw_p + 0.2
    y_ch_prawy_1 = y_kraw_p + 0.14

    # Określenie szerokości:
    if bar_prawa_rodz == 'linowa':
        bar_prawa_szer = 0.15
    else:
        bar_prawa_szer = 0.40

    if bal_prawa_rodz == 'B':
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
        if bar_prawa_rodz == 'linowa':
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
            name = f'Barieroporęcz_{bar_prawa_rodz}'
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
            name = f'Barieroporęcz_{bar_prawa_rodz}'
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
        if bal_prawa_rodz == 'B':
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
        y_wym_l = y_bar_lewa + bar_lewa_rodz + 0.19
    else:
        if bal_lewa_rodz == 'B':
            y_wym_l = y_bal_lewa + bal_lewa_wys + 0.165
        else:
            y_wym_l = y_bal_lewa + 1.24501667
    if bar_prawa == 'N':
        y_wym_p = y_bar_prawa + bar_prawa_rodz + 0.19
    else:
        if bal_prawa_rodz == 'B':
            y_wym_p = y_bal_prawa + bal_prawa_wys + 0.165
        else:
            y_wym_p = y_bal_prawa + 1.24501667
    y_wym = y_g + max(1.75, y_wym_l, y_wym_p)

    # Wymiary szczegółowe:
    wymiary_gora_1 = []

    # Wymiary elementów jezdni po lewej stronie niwelety:
    for index, row in pasy_lewe.iterrows():
        if index == 0:
            x1 = x_g
            y1 = y_g
        if isnan(pasy_lewe.iloc[index, 0]) or pasy_lewe.iloc[index, 0] == 0:
            break
        x2 = round(x1 - row['PL - szer'], 8)
        y2 = round(y1 + row['PL - szer'] * row['PL - spadek'] / 100, 8)
        wymiary_gora_1.append((aDouble(x2, y2, 0), aDouble(x1, y1, 0), 'pas ruchu'))
        x1 = x2
        y1 = y2

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
        if bal_lewa_rodz == 'B':
            x_bal_wym = x_bal_lewa + 0.04
            y_bal_wym = y_bal_lewa + bal_lewa_wys - 0.01
        else:
            x_bal_wym = x_bal_lewa + 0.175
            y_bal_wym = round(y_bal_lewa + 1.07001667, 12)
        if bar_lewa_rodz == 'linowa':
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
        if bal_prawa_rodz == 'B':
            x_bal_wym = x_bal_prawa - 0.04
            y_bal_wym = y_bal_prawa + bal_prawa_wys - 0.01
        else:
            x_bal_wym = x_bal_prawa - 0.175
            y_bal_wym = round(y_bal_prawa + 1.07001667, 12)
        if bar_prawa_rodz == 'linowa':
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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - przekrój uzytkowy narysowany!')

    return [y_wym, wymiary_gora_1, wymiary_gora_2, wymiary_gora_3, wymiary_gora_4,
            x_lewy_skr, y_lewy_skr, x_prawy_skr, y_prawy_skr, koty, pow_gorna, os]


@speed_test
def rysowanie_konstrukcja_belkowy(file, sheet, pow_gorna):
    # Pobranie danych:
    konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52], sheet_name=sheet)
    h_wsp_kr = konstrukcja.iloc[0, 0]
    h_wsp_zam = konstrukcja.iloc[0, 1]
    h_pl = konstrukcja.iloc[0, 2]
    h_pl_zam = konstrukcja.iloc[0, 3]
    szer_zam = konstrukcja.iloc[0, 4]
    h_dzw = konstrukcja.iloc[0, 5]
    b_dzw = konstrukcja.iloc[0, 6]
    n_dzw = konstrukcja.iloc[0, 7]
    roz_dzw = konstrukcja.iloc[0, 8]
    skos_dzw = konstrukcja.iloc[0, 9]

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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - ustrój betonowy, belkowy narysowany!')


@speed_test
def rysowanie_konstrukcja_skrzynkowy(y_g, file, sheet, pow_gorna):

    # Pobranie danych:
    konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59,
                                               60, 61, 62, 63], sheet_name=sheet)

    h_wsp_kr = konstrukcja.iloc[0, 0]
    h_wsp_zam = konstrukcja.iloc[0, 1]
    h_skrz = konstrukcja.iloc[0, 2]
    b_skrz = konstrukcja.iloc[0, 3]
    b_skos_l = konstrukcja.iloc[0, 4]
    b_skos_p = konstrukcja.iloc[0, 5]
    h_plg = konstrukcja.iloc[0, 6]
    h_plg_zam_l = konstrukcja.iloc[0, 7]
    b1_plg_zam_l = konstrukcja.iloc[0, 8]
    b2_plg_zam_l = konstrukcja.iloc[0, 9]
    h_plg_zam_p = konstrukcja.iloc[0, 10]
    b1_plg_zam_p = konstrukcja.iloc[0, 11]
    b2_plg_zam_p = konstrukcja.iloc[0, 12]
    h_pld = konstrukcja.iloc[0, 13]
    h_pld_zam_l = konstrukcja.iloc[0, 14]
    b1_pld_zam_l = konstrukcja.iloc[0, 15]
    b2_pld_zam_l = konstrukcja.iloc[0, 16]
    h_pld_zam_p = konstrukcja.iloc[0, 17]
    b1_pld_zam_p = konstrukcja.iloc[0, 18]
    b2_pld_zam_p = konstrukcja.iloc[0, 19]
    t_sr = konstrukcja.iloc[0, 20]

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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - ustrój betonowy, skrzynkowy narysowany!')


@speed_test
def rysowanie_konstrukcja_belki_T(file, sheet, pow_gorna):
    # Pobranie danych:
    konstrukcja = pd.read_excel(file, usecols=[43, 44], sheet_name=sheet)
    h_pl = konstrukcja.iloc[0, 0]
    bel_rodz = konstrukcja.iloc[0, 1]

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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - ustrój z belek T narysowany!')


@speed_test
def rysowanie_konstrukcja_plytowy(file, sheet, pow_gorna):
    # Pobieranie danych:
    konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49], sheet_name=sheet)
    h_wsp_kr = konstrukcja.iloc[0, 0]
    h_wsp_zam = konstrukcja.iloc[0, 1]
    h_pl_pl = konstrukcja.iloc[0, 2]
    h_pl = konstrukcja.iloc[0, 3]
    b_pl = konstrukcja.iloc[0, 4]
    b_pl_skos_l = konstrukcja.iloc[0, 5]
    b_pl_skos_p = konstrukcja.iloc[0, 6]

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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - ustrój betonowy, płytowy narysowany!')


@speed_test
def rysowanie_konstrukcja_zespolony(file, sheet, pow_gorna):
    konstrukcja = pd.read_excel(file, usecols=[43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56], sheet_name=sheet)
    h_wsp = konstrukcja.iloc[0, 0]
    h_pl = konstrukcja.iloc[0, 1]
    skos_pl = konstrukcja.iloc[0, 2]
    skos_h = konstrukcja.iloc[0, 3]
    h_dzw = konstrukcja.iloc[0, 4]
    n_dzw = konstrukcja.iloc[0, 5]
    roz_dzw = konstrukcja.iloc[0, 6]
    b_pg = round(konstrukcja.iloc[0, 7] / 1000, 8)
    t_pg = round(konstrukcja.iloc[0, 8] / 1000, 8)
    b_pd = round(konstrukcja.iloc[0, 9] / 1000, 8)
    t_pd = round(konstrukcja.iloc[0, 10] / 1000, 8)
    t_sr = round(konstrukcja.iloc[0, 11] / 1000, 8)
    z_od_g = round(konstrukcja.iloc[0, 12] / 1000, 8)
    z_od_d = round(konstrukcja.iloc[0, 13] / 1000, 8)

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

    if len(sheet.split("_")) == 2:
        obiekt = sheet.split("_")[1]
    else:
        obiekt = f'{sheet.split("_")[1]}_{sheet.split("_")[2]}'
    print(f'Obiekt {obiekt} - ustrój zespolony narysowany!')


@speed_test
def opis(opis_gora, file, sheet):
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
    obiekt = sheet.split("_")[1]
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
    delta_y = pd.read_excel(file, usecols=[0], sheet_name=sheet).iloc[0, 0]

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

    obiekt = sheet.split("_")[1]
    print(f'Obiekt {obiekt} - opis gotowy!')
