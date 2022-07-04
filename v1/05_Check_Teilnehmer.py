# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 09:41:35 2020

@author: ufriess

Required packages:
    pytzbar   - pip install pyzbar
    fitz      - pip install PyMuPDF
    PIL       - pip install pillow
    pdf2image - pip install pdf2image

INPUT: 
    klausurteilnehmer.xls als Export aus der Übungsgruppendatenbank 
    mit folgenden Einträgen:
        - Nachname
        - Vorname
        - UID
        - Matrikelnummer
        - Gruppe
"""

import pandas as pd
import numpy as np
import warnings


def add_teilnehmer_info(klausur_map, teilnehmer):
    
    for c in teilnehmer.columns:
        klausur_map[c] = ''
        
    idx = klausur_map.index.names
    
    klausur_map = klausur_map.reset_index().set_index('UID')
    for ig, g in klausur_map.groupby(level = 'UID'):
        for c in teilnehmer.columns:
            klausur_map.loc[klausur_map.index == ig, c] = teilnehmer.loc[ig, c]

    klausur_map = klausur_map.reset_index().set_index(idx)

    return klausur_map



######### Main Program
warnings.filterwarnings("ignore")


# Prüfen, ob PDF aller Teilnehmer da
check_teilnahme = True 

num_aufgaben = 12
seiten_pro_aufgabe = [2 if a < 11 else 6 for a in range(num_aufgaben)]

klausur_map = pd.read_pickle('klausur_map.pkl')

# Manuelle Zuweisung nicht erkannter Seiten


#%% Lade Klausurteilnehmer
usecols = ['Nachname', 'Vorname', 'UID', 'Matrikel']
if check_teilnahme:
    usecols = usecols + ['Teilgenommen']
    
teilnehmer = pd.read_excel('klausurteilnehmer.xls', header = 0,
                           usecols = usecols)

if teilnehmer.iloc[0].Nachname is np.NaN: # first row is empty
    teilnehmer.drop(0, inplace = True)

teilnehmer.UID = np.nan_to_num(teilnehmer.UID).astype(int)
teilnehmer.Matrikel = np.nan_to_num(teilnehmer.Matrikel).astype(int)
#teilnehmer.Gruppe = teilnehmer.Gruppe.astype(int)

if 'Teilgenommen' in teilnehmer.columns:
    teilnehmer.Teilgenommen = teilnehmer.Teilgenommen.astype(bool)

teilnehmer.set_index('UID', inplace = True)

#%% Liste mit Daten für Seiten, auf denen der QR-Code nicht erkannt wurde
# missing_qr = pd.read_csv('missing_qr.dat', sep = '\t', comment = '#',
#                          index_col = ['UID','Aufgabe','Seite','Gruppe'])
missing_qr = pd.read_excel('missing_qr.xlsx')
missing_qr.set_index(['UID','Aufgabe','Seite','Gruppe'], inplace = True)

klausur_map = pd.concat([klausur_map, missing_qr])

klausur_map = klausur_map.loc[~klausur_map.index.duplicated(keep='first')].sort_index()
#%% Check for duplicates in klausur_map
# duplicates = klausur_map[klausur_map.index.duplicated()]
# if len(duplicates) > 0:
#     raise Exception('Doppelte Einträge gefunden! Möglicherweise Fehler in der Datei "missing_qr.dat"')

#%% Check for missing participants

print('Prüfe auf Vollständigkeit')

error = False

for uid in teilnehmer.index:
    if check_teilnahme:
        if teilnehmer.loc[uid].Teilgenommen and (not uid in klausur_map.index.get_level_values('UID')):
            print('Keine Klausur für', uid, teilnehmer.loc[uid].Nachname, teilnehmer.loc[uid].Vorname)
            error = True

        if (not teilnehmer.loc[uid].Teilgenommen) and (uid in klausur_map.index.get_level_values('UID')):
            print('Unerwartete Klausur für', uid, teilnehmer.loc[uid].Nachname, teilnehmer.loc[uid].Vorname)
            error = True
    else:
        if not uid in klausur_map.index.get_level_values('UID'):
            print('Keine Klausur für', uid, teilnehmer.loc[uid].Nachname, teilnehmer.loc[uid].Vorname)
            error = True
        
for uid, stud in klausur_map.groupby(level = 0):
    for a in range(num_aufgaben-1):
        for s in range(seiten_pro_aufgabe[a]):
            try:
                assert len(stud.loc(axis = 0)[:,a,s+1,:]) > 0
            except:
                print('Aufgabe', a, 'Seite', s+1, 'fehlt für', uid, teilnehmer.loc[uid].Nachname, teilnehmer.loc[uid].Vorname)
                error = True

if not error:
    print('Alles Okay :-)')        
            
#%% Füge anhand der UID Nachname, Vorname, Uni-ID der klausur_map zu und sortiere alphabetisch    
print('Schreibe Datenbank')

klausur_map = add_teilnehmer_info(klausur_map, teilnehmer)    
klausur_map.to_pickle('klausur_map_teilnehmer.pkl')
