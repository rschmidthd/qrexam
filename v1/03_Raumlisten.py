# -*- coding: utf-8 -*-
"""
Created on Mon Jul 27 13:52:41 2020

@author: ufriess
"""

import pandas as pd
import numpy as np

# Lese Teilnehmer
teilnehmer = pd.read_excel('klausurteilnehmer.xls', header = 0,)
                            # usecols = ['Nachname', 'Vorname', 'UID', 'Matrikel', 
                            #            'Gruppe', 'Raum', 'Sitzplatz'])

if teilnehmer.iloc[0].Nachname is np.NaN: # first row is empty
    teilnehmer.drop(0, inplace = True)
    
teilnehmer.UID = np.nan_to_num(teilnehmer.UID).astype(int)
teilnehmer.Matrikel = np.nan_to_num(teilnehmer.Matrikel).astype(int)
teilnehmer.Gruppe = teilnehmer.Gruppe.astype(int)

teilnehmer.sort_values(by = ['Raum', 'Sitzplatz'], inplace = True)

writer = pd.ExcelWriter('raumlisten.xls')

for i, raum in teilnehmer.groupby('Raum'):
    raum.loc(axis = 1)['Sitzplatz', 'Raum', 'Nachname', 'Vorname', 'Matrikel'].to_excel(
        writer, sheet_name = i, index = False)
    print(i, raum.Sitzplatz.values)

writer.close()