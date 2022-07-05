# -*- coding: utf-8 -*-
"""

@author: ufriess

Adapted Robert Schmidt Jan-May 2022


INPUT: 
    pdf of exam
    number of exams: NumKlausuren
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


def AddKlausur(output, num, add_empty_page = False):

    def deckblatt(can):
    
        # enter exam number on front piece
        can.setFontSize(12)
        can.drawString(230, 530, '{:03d}'.format(num))
    
    def fusszeile(can, num, aufgabe, seite):
    
        # Text
        can.setFontSize(10)
        can.drawString(120, 25, 'Klausur Nr {:03d}, A{:02d}/{:02d}'.format(
            num, aufgabe, seite))  # num=Nummer der Klausur
        
        # QR code
        qr = qrcode.QRCode(
         #     box_size=200,
         #     border=4,
         )
        #                    UID    Gruppe     Aufgabe    Seite                           
        qr.add_data('{:03d},{:02d},{:02d}'.format(
            num, aufgabe, seite))
        qr.make(fit=True)
        qr = qr.make_image(fill_color="black", back_color="white")   
        
        io_img = io.BytesIO()
        qr.save(io_img, 'PNG')
        
        reportlab_io_img = ImageReader(io_img)
        
        can.drawImage(reportlab_io_img, 30, 15, width = 60, 
                      preserveAspectRatio = True,
                      anchor = 'sw')

        # Horizontale Linie
        can.setLineWidth(0.5)
        can.line(30, 80, 560, 80)
        can.setFontSize(10)
        can.drawString(120, 40, 'Diesen Bereich bitte nicht beschriften')
        
    
    a = 0
    s = 1
    
    outpagenum = 1
    
    for pagenum, page in enumerate(Klausur.pages):
        
        qrpage = io.BytesIO()
        
        can = canvas.Canvas(qrpage, pagesize=A4)
    
        if pagenum == 0:
        #    if data.Gruppe >0:
            deckblatt(can)
            
        if a==9 and s==1:          # Zusätzblätter haben Seitenzahl >=3 um
            s=s+2                  # zugeordnet werden zu können
        fusszeile(can, num, a, s)
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
    
        if a <= NumAufgaben:  # for regular problems increase count
            a += 1
        else:         # for additional sheets raise page number by two
            s +=2

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
outpath = 'tray'

infile = 'klausur_demo.pdf'

# Nur eine einzige Datei mit allen Klausuren?
# Falls SingleFile = False, dann einzelne Ausgabedatei für jeden Raum.
# SingleFile = False empfohlen für Hauptklausur
SingleFile = False

# Ausgabedatei, falls SingleFile = True
outfile = 'Exams_PEP1_year.pdf'



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

NumAufgaben =8 
NumKlausuren=10


# numbers avoided in planes and hotels
avoid_numbers=(13,14,17,113,114,117,213,214,217,313,314,317,413,414,417)

for i in range(NumKlausuren):

    if i%30==0 and i>0:
        print('{:02d}%'.format(int(100*i/NumKlausuren)))
    if i+1 in avoid_numbers:
        continue
    
    output = PdfFileWriter()
    AddKlausur(output, i+1, add_empty_page = add_empty_page)
            
    # finally, write "output" to a real file
    outputStream = open(os.path.join(outpath, os.path.basename(infile).split('.')[0] + '_{:03d}'.format(i+1) + '.pdf'), 'wb')
    output.write(outputStream)
    outputStream.close()

