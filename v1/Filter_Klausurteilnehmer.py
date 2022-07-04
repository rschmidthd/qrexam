# -*- coding: utf-8 -*-
"""
Created on Fri Jul 17 09:39:54 2020

@author: ufriess
"""
import pandas as pd
import numpy as np


def read_teilnehmer(filename):

    teilnehmer = pd.read_excel(filename, header = 1) 
    
    if teilnehmer.iloc[0].Nachname is np.NaN: # first row is empty
        teilnehmer.drop(0, inplace = True)
        
    teilnehmer.Matrikel = np.nan_to_num(teilnehmer.Matrikel).astype(int)
    teilnehmer.UID = np.nan_to_num(teilnehmer.UID).astype(int)
#    teilnehmer['Summe Uebungen'] = np.nan_to_num(teilnehmer['Summe Uebungen']).astype(int)
    teilnehmer.Gruppe = teilnehmer.Gruppe.astype(int)
    
    return teilnehmer
    

teilnehmer = read_teilnehmer('./uebungsteilnehmer.xls')
teilnehmer.set_index('E-mail', inplace = True)
#%%
with open('email_teilnehmer.txt', 'r') as f:
    mails = f.readline().split(';')
    
mails = [m.strip(" ").strip("'") for m in mails]

teilnehmer = teilnehmer.loc[mails]

teilnehmer.to_excel('klausurteilnehmer.xls')