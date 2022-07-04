# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 16:14:22 2020

@author: ufriess

INPUT: 
    klausurteilnehmer.xls exportiert aus der Übungsgruppendatenbank 
    mit folgenden Einträgen (Reihenfolge ist wichtig):
        - Nachname
        - Vorname
        - UID
        - Matrikelnummer
        - Gruppe
"""

import pandas as pd
import numpy as np
import io
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader


def AddKlausur(output, data, add_empty_page = False):

    def deckblatt(can):
    
        x = deckblatt_coord[0]
        y = deckblatt_coord[1]
        can.drawString(x, y, data.Nachname)
        y -= deckblatt_delta
        can.drawString(x, y, data.Vorname)
        y -= deckblatt_delta
        can.drawString(x, y, '{:d}'.format(data.Matrikel))
        y -= deckblatt_delta
        can.drawString(x, y, '{:02d}'.format(data.Gruppe))
        y -= deckblatt_delta
        can.drawString(x, y, data.Raum)
        y -= deckblatt_delta
        can.drawString(x, y, '{:02d}'.format(data.Sitzplatz))
    
    def fusszeile(can, aufgabe, seite):
    
        # Text
        can.setFontSize(12)
        can.drawString(500, 50, '{:d}'.format(outpagenum))

        can.setFontSize(10)
        can.drawString(120, 50, '{}, {}, {:d}, G{:d}, {}-{:d}, A{:d}/{:d}'.format(
            data.Nachname, data.Vorname, data.UID, data.Gruppe, 
            data.Raum, data.Sitzplatz, aufgabe, seite))
        
        # QR code
        qr = qrcode.QRCode(
         #     box_size=200,
         #     border=4,
         )
        #                    UID    Gruppe     Aufgabe    Seite                           
        qr.add_data('{:d},{:d},{:d},{:d}'.format(
            data.UID, data.Gruppe, aufgabe, seite))
        qr.make(fit=True)
        qr = qr.make_image(fill_color="black", back_color="white")   
        
        io_img = io.BytesIO()
        qr.save(io_img, 'PNG')
        
        reportlab_io_img = ImageReader(io_img)
        
        can.drawImage(reportlab_io_img, 30, 25, width = 60, 
                      preserveAspectRatio = True,
                      anchor = 'sw')

        # Horizontale Linie
        can.setLineWidth(0.5)
        can.line(30, 80, 560, 80)
        can.setFontSize(10)
        can.drawString(120, 70, 'Diesen Bereich bitte nicht beschriften')
        
    
    a = 0
    s = 1
    
    outpagenum = 1
    
    for pagenum, page in enumerate(Klausur.pages):
        
        qrpage = io.BytesIO()
        
        can = canvas.Canvas(qrpage, pagesize=A4)
    
        if pagenum == 0:
            if data.Gruppe >0:
                deckblatt(can)
            
        fusszeile(can, a, s)
        can.save()
        
        # move to the beginning of the StringIO buffer
        qrpage.seek(0)
        qr_pdf = PdfFileReader(qrpage)
        
        # add QR code
        # page.mergePage(qr_pdf.getPage(0))
        # output.addPage(page)
        qr_pdf.getPage(0).mergePage(page)
        output.addPage(qr_pdf.getPage(0))
        outpagenum += 1
    
        s += 1
        if (s > 2) and (a <= NumAufgaben):
            a += 1
            s = 1

        # Erzeuge leere Seite - außer für Zusatzblätter
        if a <= NumAufgaben:
            if add_empty_page:
                qrpage = io.BytesIO()
            
                can = canvas.Canvas(qrpage, pagesize=A4)
                fusszeile(can, a, s)
                can.save()
                
                # move to the beginning of the StringIO buffer
                qrpage.seek(0)
                qr_pdf = PdfFileReader(qrpage)
            
                output.addPage(qr_pdf.getPage(0))
                outpagenum +=1

                s += 1
                if (s > 2) and (a <= NumAufgaben):
                    a += 1
                    s = 1
        

#### Main Program
outpath = 'klausuren_druckvorlage'

infile = '../Nachklausur.pdf'
teilnehmerfile = 'klausurteilnehmer.xls'

# Nur eine einzige Datei mit allen Klausuren?
# Falls SingleFile = False, dann einzelne Ausgabedatei für jeden Raum.
# SingleFile = False empfohlen für Hauptklausur
SingleFile = False 

# Ausgabedatei, falls SingleFile = True
outfile = 'Nachklausuren_PEP1_WS1920.pdf'


# Anzahl Leere Klausuren 
# Für Studierende, die nicht auf der Liste sind oder im falschen Raum
n_leer = 2


# infile = '../Nachklausur_large.pdf'
# teilnehmerfile = 'klausurteilnehmer_largefont.xls'
# outfile = 'Nachklausuren_PEP1_WS1920_largefont.pdf'
# # Leere Klausuren
# n_leer = 0

add_empty_page = False


try:
    os.makedirs(outpath)
except OSError:
    pass
    
# Obere linke Koordinate für Teilnehmerinfo auf Deckblatt
deckblatt_coord = [300, 705] 
# Zeilenabstand für Teilnehmerinfo auf Deckblatt
deckblatt_delta = 23

Klausur = PdfFileReader(open(infile, 'rb'))
NumAufgaben = 10

#%%
teilnehmer = pd.read_excel(teilnehmerfile, header = 0,
                            usecols = ['Nachname', 'Vorname', 'UID', 'Matrikel', 
                                       'Gruppe', 'Raum', 'Sitzplatz'])

if teilnehmer.iloc[0].Nachname is np.NaN: # first row is empty
    teilnehmer.drop(0, inplace = True)
    
teilnehmer.UID = np.nan_to_num(teilnehmer.UID).astype(int)
teilnehmer.Matrikel = np.nan_to_num(teilnehmer.Matrikel).astype(int)
teilnehmer.Gruppe = teilnehmer.Gruppe.astype(int)
teilnehmer.Sitzplatz = teilnehmer.Sitzplatz.astype(int)

teilnehmer.sort_values(by = ['Raum', 'Sitzplatz'], inplace = True)

#teilnehmer = teilnehmer.iloc[-1:]

for i in range(n_leer):
    teilnehmer = teilnehmer.append(pd.Series({'Raum': 'Extra', 
                                              'Sitzplatz':i, 
                                              'UID': 1000000 + i,
                                              'Matrikel': 0,
                                              'Nachname': '',
                                              'Vorname': '',
                                              'Gruppe': 0}), ignore_index = True)

#teilnehmer = teilnehmer[teilnehmer.Raum == 'Extra']
#%% Getrennte Datei für jeden Raum

if SingleFile:
    print('Erstelle PDF')
    
    output = PdfFileWriter()
    
    for i, s in teilnehmer.iterrows():
        AddKlausur(output, s, add_empty_page = add_empty_page)
        
    # finally, write "output" to a real file
    outputStream = open(outpath + '\\' + outfile, 'wb')
    output.write(outputStream)
    outputStream.close()
else:
    for ig, g in teilnehmer.groupby('Raum'):
        print('Erstelle PDF für Raum {}'.format(ig))
    
        output = PdfFileWriter()
    
        for i, s in g.iterrows():
            AddKlausur(output, s, add_empty_page = add_empty_page)
            
        # finally, write "output" to a real file
        outputStream = open(os.path.join(outpath, os.path.basename(infile).split('.')[0] + '_{}'.format(ig) + '.pdf'), 'wb')
        output.write(outputStream)
        outputStream.close()

