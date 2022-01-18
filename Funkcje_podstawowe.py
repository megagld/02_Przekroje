from pyautocad import aDouble, Autocad
import win32com.client
import pythoncom
from time import time
from functools import wraps
import pandas as pd


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

def ustal_pasy_ruchu(tab):
    #ustala szerokości pasów ruchu i kierunki oraz ustala tabele pd
    pasy=tab.iloc[0,0].split('+')
    kierunki=tab.iloc[0,3].split('+')
    kierunki=kierunki*(len(pasy)-len(kierunki)+1)

    headers=tab.columns
    tab=pd.DataFrame([tab.iloc[0,].tolist()]*(len(pasy)))
    tab.columns=headers

    szer_h=headers[0]
    kier_h=headers[3]

    for i,j in enumerate(zip(pasy,kierunki)):
        szer=j[0]
        kier=j[1]
        tab.loc[i,szer_h]=float(szer)
        tab.loc[i,kier_h]=kier
    
    return tab
