#not/usr/bin/env python
# -*- coding: utf-8 -*-
import pandas as pd
from icecream import ic
ic.configureOutput(prefix='🐍 debug| -> ')
try:
    data = pd.read_excel('data.xls')
except (FileNotFoundError, NameError, KeyError) as e:
    ic("🛑 erreur(s)")
    ic(e)
else:
    clean = data.pivot_table(index="Bâtiment", columns="Code surface", values='Surface calculée')
    clean=clean.reset_index()
    ic(clean.head(5))
    data.drop(['Code surface', 'Surface calculée'], axis=1, inplace=True)
    data.drop_duplicates(inplace=True)
    clean= clean.set_index('Bâtiment').join(data.set_index('Bâtiment'))
    clean.rename(columns={"Nombre de postes de travail":'pdt'}, inplace=True)
    clean.drop(['SHOB', 'SHON'], axis=1, inplace=True)
    clean= clean.assign(sun_sub= lambda x: x.SUN / x.SUB)
    clean= clean.assign(sun_pdt= lambda x: x.SUN / x.pdt )
    clean.rename(columns={"sun_pdt":'SUN/pdt', "sun_sub":'SUN/SUB'}, inplace=True)
    clean= clean.reset_index()
    clean= clean.assign(bat_de_bureau='')
    for i in range(len(clean)):
        if clean.loc[i,'SUN/SUB'] >= .5:
            clean.loc[i,'bat_de_bureau'] = 'oui'
    ic(clean["Site"].unique().size)
    ic(clean["Bâtiment"].unique().size)
    clean.to_excel('data_par_bat.xlsx', index = None, header=True)        
    clean =clean.query('bat_de_bureau == "oui"')
    clean = clean.reindex(columns=['Département','Commune', '0B - Numéro Chorus', 'Site', 'Bâtiment', '0B - Numéro Chorus.1', 'Code Postal', 'SDP', 'SUB','SUN', "SUN/SUB", 'pdt', 'SUN/pdt' ])
    clean.to_excel('ratio_par_bat.xlsx', index = None, header=True)
