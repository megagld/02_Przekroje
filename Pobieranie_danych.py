from tarfile import LENGTH_NAME
import pandas as pd
from pathlib import Path
from Funkcje_podstawowe import *

def pobierz_dane(file,sheet,nr_przekroju,s_def_name):
    # Pobranie danych:
    global tabela,tom,obiekt,typ,delta_y,delta_x,pasy_lewe,pasy_prawe,awaryjny_lewy,awaryjny_prawy,opaska_lewa,opaska_prawa,chodnik_lewy,chodnik_prawy,bariery_lewa,bariery_prawa,zalamanie_lewe,zalamanie_prawe,konstrukcja,ramka

    dane = pd.read_excel(file, usecols=[nr_przekroju], sheet_name=sheet).T.iloc[:4,:110]
    ozn = pd.read_excel(file, usecols=[1], sheet_name=sheet).T.iloc[:,:110]
    headers=[str(i) for i in ozn.iloc[0,:]]
    dane.columns=headers

    # global ramka
    dane_def = pd.read_excel(file, usecols=[4], sheet_name=s_def_name).T.iloc[:4,:110]
    ozn_def = pd.read_excel(file, usecols=[1], sheet_name=s_def_name).T.iloc[:,:110]
    headers_def=[str(i) for i in ozn_def.iloc[0,:]]
    dane_def.columns=headers_def

    # zmiana '-' na zera
    for (colname,colval) in dane.iteritems():
        if colval.values=='-':
            dane[colname]=0

    dane.index=[0]
    tabela=dane
    tom=tabela['Tom'].to_frame().iloc[0, 0]
    obiekt=tabela['Obiekt'].to_frame().iloc[0, 0]
    typ=tabela['Typ obiektu'].to_frame().iloc[0, 0]
    delta_y=tabela['Δ niw'].to_frame().iloc[0, 0]
    delta_x=tabela['Δ niw_poz'].to_frame().iloc[0, 0]
    pasy_lewe = ustal_pasy_ruchu(tabela.loc[:,'PL - szer':'PL - kier'])
    pasy_prawe = ustal_pasy_ruchu(tabela.loc[:,'PP - szer': 'PP - kier'])
    awaryjny_lewy = tabela.loc[:,'PAL - szer':'PAL - spadek']
    awaryjny_prawy = tabela.loc[:,'PAP - szer':'PAP - spadek']
    opaska_lewa = tabela.loc[:,'OL - szer':'OL - spadek']
    opaska_prawa = tabela.loc[:,'OP - szer': 'OP - spadek']
    chodnik_lewy = tabela.loc[:,'CL - szer CH':'LL - T/N'] 
    chodnik_prawy = tabela.loc[:,'CP - szer CH': 'LP - T/N']
    bariery_lewa = tabela.loc[:,'BL - T/N': 'B/E L - wys']
    bariery_prawa = tabela.loc[:,'BP - T/N':'B/E P - wys']
    zalamanie_lewe = tabela.loc[:,'ZL T/N':'ZL - x']
    zalamanie_prawe = tabela.loc[:,'ZP T/N':'ZP - x']

    #dane domyślne
    dane_def.index=[0]
    tabela_def=dane_def
    ramka=tabela_def['ramka'].to_frame().iloc[0, 0]

    # # Pobranie danych:
    konstrukcja =tabela.loc[:,'B_WSP - h':'Z_Ż - od d']