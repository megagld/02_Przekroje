from pathlib import Path
import pandas as pd

# Plik z danymi:
file = '{}{}'.format(Path(__file__).parent,'''\Dane.xlsx''')
# file = 'c:\_magazyn\_Python\_skrypty\02_Przekroje\transpozycja i nagówki\343 - testy pandy - transpozycja.p'
print(file)
# Nazwy arkuszy:
sheet_names = pd.ExcelFile(file).sheet_names
print(sheet_names)

sheet=sheet_names[0]

dane = pd.read_excel(file, usecols=[3,4,5,6], sheet_name=sheet).T.iloc[:4,:57]
ozn = pd.read_excel(file, usecols=[9], sheet_name=sheet).T.iloc[:,:57]
headers=[str(i) for i in ozn.iloc[0,:]]
dane.columns=headers
dane.index=[0,1,2,3]
tabela=dane

# print(tabela)
q=tabela['PL - szer']
print(q)
# print(tabela.iloc[1,1])

# print(headers)

# dane na sztywno:
# ['Tom', 'Obiekt', 'Typ obiektu', 'Δ niw', 'Δ niw_poz', 'PL - szer', 'PL - spadek', 'PL - kier rodz', 'PL - kier', 'PP - szer', 'PP - spadek', 'PP - kier rodz', 'PP - kier', 'PAL - szer', 'PAL - spadek', 'PAP - szer', 'PAP - spadek', 'OL - szer', 'OL - spadek', 'OP - szer', 'OP - spadek', 'CL - szer CH', 'CL - szer ŚR', 'CL - szer CPR', 'CL - spadek', 'CL - deska', 'LL - T/N', 'CP - szer CH', 'CP - szer ŚR', 'CP - szer CPR', 'CP - spadek', 'CP - deska', 'LP - T/N', 'BL - T/N', 'BL - rodzaj', 'BL - opaska', 'B/E L - rodz', 'B/E L - wys', 'BP - T/N', 'BP - rodzaj', 'BP - opaska', 'B/E P - rodz', 'B/E P - wys', 'ZL T/N', 'ZL - x', 'ZP T/N', 'ZP - x', 'WSP - h', 'WSP - h zam', 'PL - h', 'PL - h zam', 'PL - szer zam', 'DZW - h', 'DZW - b', 'DZW - n', 'DZW - roz', 'DZW - skos']