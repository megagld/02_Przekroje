import pandas as pd
from pathlib import Path

def pobierz_dane(file,sheet):
    # Pobranie danych:
    global tabela,tom,obiekt,typ,delta_y,delta_x,pasy_lewe,pasy_prawe,awaryjny_lewy,awaryjny_prawy,opaska_lewa,opaska_prawa,chodnik_lewy,chodnik_prawy,bariery_lewa,bariery_prawa,zalamanie_lewe,zalamanie_prawe,konstrukcja
    
    dane = pd.read_excel(file, usecols=[11,12,13,14], sheet_name=sheet).T.iloc[:4,:57]
    # do usunięcia powyższe - !!!
    # dane = pd.read_excel(file, usecols=[3,4,5,6], sheet_name=sheet).T.iloc[:4,:57]
    ozn = pd.read_excel(file, usecols=[9], sheet_name=sheet).T.iloc[:,:57]
    headers=[str(i) for i in ozn.iloc[0,:]]
    dane.columns=headers
    dane.index=[0,1,2,3]
    tabela=dane

    tom=tabela['Tom'].to_frame().iloc[0, 0]
    obiekt=tabela['Obiekt'].to_frame().iloc[0, 0]
    typ=tabela['Typ obiektu'].to_frame().iloc[0, 0]
    delta_y=tabela['Δ niw'].to_frame().iloc[0, 0]
    delta_x=tabela['Δ niw_poz'].to_frame().iloc[0, 0]
    pasy_lewe = tabela.loc[:,'PL - szer':'PL - kier']
    pasy_prawe = tabela.loc[:,'PP - szer': 'PP - kier']
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

    # # Pobranie danych:
    konstrukcja =tabela.loc[:,'WSP - h':'DZW - skos']
