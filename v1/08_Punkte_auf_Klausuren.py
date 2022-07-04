# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 16:14:22 2020

@author: ufriess

"""

import pandas as pd
import numpy as np
import io
import os
import random
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait, landscape


def Addpunkte(infile, outfile, punkte, aufgaben_seiten):

    
    punkte = np.nan_to_num(punkte)
      
    try:
        os.makedirs(os.path.dirname(outfile))
    except OSError:
        pass
    
    
    Klausur = PdfFileReader(open(infile, 'rb'))
    output = PdfFileWriter()

    for pagenum, page in enumerate(Klausur.pages):
        p_punkte = io.BytesIO()    
        can = canvas.Canvas(p_punkte, pagesize=portrait(A4), invariant = 1)
        # can.saveState()
        # can.rotate(90)
        merge = False
        
        if pagenum == 1:
            # Punktetabelle auf Zweites Blatt
            
            x = punkte_coord[0]
            y = punkte_coord[1]
            
            can.rect(x,y,60,25)
            can.rect(x,y+25,60,25)
            can.drawString(x+5,y+5, 'Punkte')
            can.drawString(x+5,y+30, 'Aufgabe')
            
            xpos = x + 60
            for i, n in enumerate(punkte):
                can.rect(xpos, y, punkte_delta, 25)
                can.rect(xpos, y+25, punkte_delta, 25)

                can.drawString(xpos + 5, y+5, str(n))
                can.drawString(xpos + 5, y+30, str(i+1))
                
                xpos = xpos + punkte_delta

            # Punktesumme                
            can.rect(xpos, y, 60, 25)
            can.rect(xpos, y+25, 60, 25)
            can.setFont('Helvetica-Bold', 12)    
            can.drawString(xpos + 5, y+5, str(sum(punkte)))
            can.drawString(xpos + 5, y+30, 'Summe', direction = 90)
            

            merge = True
            # move to the beginning of the StringIO buffer
            
        # if pagenum in a_s:
        #     p = punkte[np.where(a_s == pagenum)[0][0]]
        #     can.setFontSize(16)
        #     can.drawString(apunkte_coord[0], apunkte_coord[1], str(p))
        #     merge = True

       
        if merge:
            # can.restoreState()
            can.save()

            p_punkte.seek(0)
            pdf_punkte = PdfFileReader(p_punkte)
 
            page = pdf_punkte.getPage(0)
#            page.mergePage(pdf_punkte.getPage(0))

        output.addPage(page)

    # finally, write "output" to a real file
    outputStream = open(outfile, 'wb')
    output.write(outputStream)
    outputStream.close()

#### Main Program
inpath = 'klausuren_einzeln_korrigiert'
outpath = 'klausuren_einzeln_korrigiert_mit_punkten'

# generiere zufällige Punkte für Testbetrieb?
# random_punkte = False

# schreibe Punkte für jede Aufgabe?
punkte_je_aufgabe = False

try:
    os.makedirs(outpath)
except OSError:
    pass
    
# Koordinate für die punktetabelle
punkte_coord = [75, 130] 
# Abstand der Spalten der punktetabelle
punkte_delta = 32

# koordinaten für die punkte für jede aufgabe
apunkte_coord = [500, 785]

#%% Lese punkte
punkte = pd.read_excel('klausurpunkte.xls', header = 3).set_index('ID').fillna(0)
aufgaben = [c for c in punkte.columns if c.startswith('Aufgabe')]

# if random_punkte:
#     punkte.loc(axis=1)[aufgaben] = punkte.loc(axis=1)[aufgaben].applymap(lambda x: random.randint(0,10))    

# for i, r in punkte.iterrows():
#     if r[aufgaben].isna().any():
#         print('Es fehlen punkte für', r.UID, r.Vorname, r.Nachname)
 
punkte['Summe'] = punkte.loc(axis = 1)[aufgaben].sum(axis = 1)       

# aufgaben = aufgaben + ['Summe']

#%%
klausur_map = pd.read_pickle('klausur_map_teilnehmer.pkl')

#klausur_map = klausur_map[:24] # Test - nur eine Klausur bearbeiten

print('Schreibe Punkte auf Klausuren...')
# loop über Teilnehmer
for i, s in klausur_map.groupby(level = 'UID'):
    if punkte.loc[i, aufgaben].isna().any():
        print('Es fehlen punkte für', i, s.Vorname[0], s.Nachname[0])

    aufgaben_seiten = s.loc[i, :, 1,:].PDFSeite_klausur.loc[1:10].values
    Addpunkte(os.path.join(inpath, s.Filename_klausur.iloc[0]), 
             os.path.join(outpath, s.Filename_klausur.iloc[0]),
             punkte.loc[i, aufgaben].values, aufgaben_seiten)
print('Fertig!')
   
