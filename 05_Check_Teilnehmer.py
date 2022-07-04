# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 09:41:35 2020

@author: ufriess
Adapted by R. Schmidt Jan/Feb 2022
 
This code checks whether the exams are complete (all problems present).


Required packages:
    pytzbar   - pip install pyzbar
    fitz      - pip install PyMuPDF
    PIL       - pip install pillow
    pdf2image - pip install pdf2image

INPUT: 
    klausur_map.pkl     data base from previous step 01
"""

import pandas as pd
import numpy as np
import warnings


######### Main Program
warnings.filterwarnings("ignore")


num_aufgaben = 9
seiten_pro_aufgabe = [2 if a < 9 else 6 for a in range(num_aufgaben)]

klausur_map = pd.read_pickle('klausur_map.pkl')


print('Checking for completeness')

error = False

# this was introduced because one problem was missing from
# the pkl data base beacsue it had been correct by hand
# (rather than scanned)
checkone = True
if not checkone:
    print ('not checking for problem 1')

for num, stud in klausur_map.groupby(level = 0):
    for a in range(num_aufgaben-1):
        for s in range(seiten_pro_aufgabe[a]):
            try:
                assert len(stud.loc(axis = 0)[:,a,s+1,:]) > 0
            except:
                if (a!=1 or checkone):
                    print('Problem', a, 'page', s+1, 'missing for', num)
                error = True

