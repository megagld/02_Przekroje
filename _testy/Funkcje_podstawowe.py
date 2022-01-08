from pyautocad import aDouble, Autocad
import win32com.client
import pythoncom
from time import time
from functools import wraps

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
