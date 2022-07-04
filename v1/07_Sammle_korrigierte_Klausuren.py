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
import os
from pdf2image import convert_from_path
from PyPDF2 import PdfFileWriter, PdfFileReader
import warnings


def sort_by_gruppe_teilnehmer(klausur_map, inpath, outpath, folder_per_gruppe = False):
    
    klausur_map['PDFSeite_klausur'] = -1
    klausur_map['Filename_klausur'] = ''    
    
    try:
        os.makedirs(outpath)
    except OSError:
        pass
    
    # Create PDFs in dictionary
    PDFs = {}
    for f, g in klausur_map.groupby('Filename_aufgabe'):
        PDFs[f] = PdfFileReader(os.path.join(inpath, f), strict = False)

    os.makedirs(outpath, exist_ok = True)

    for g, mg in klausur_map.groupby(level = 'Gruppe'):
        if folder_per_gruppe:
            print('Gruppe', g)
            gpath = os.path.join(outpath, 'Gruppe{:02d}'.format(g))
            try:
                os.makedirs(gpath)
            except OSError:
                pass
        
        for k, mk in mg.groupby(level = 'UID'):
            pdflen = 0
            pdf_writer = PdfFileWriter()
            Nachname = '-'.join(mk.iloc[0].Nachname.split(' ')) + '_' + '-'.join(mk.iloc[0].Vorname.split(' '))
            
            if folder_per_gruppe:
                outname = os.path.join('Gruppe{:02d}'.format(g), 'Klausur_{}_{:d}.pdf'.format(Nachname, k))
            else:
                outname = 'Klausur_{}_{:d}.pdf'.format(Nachname, k)

            for i, r in mk.iterrows():
#                print(r)
                pdflen += 1
                pdf_writer.addPage(PDFs[r.Filename_aufgabe].getPage(r.PDFSeite_aufgabe - 1)) 
                
                klausur_map.PDFSeite_klausur.loc[i] = pdflen
                klausur_map.Filename_klausur.loc[i] = outname
                
            with open(os.path.join(outpath, outname), 'wb') as out:
                try:
                    pdf_writer.write(out)
                except:
                    print('Error writing', outname)
                    raise
                
    return klausur_map

def sort_by_gruppe_teilnehmer_jpg(klausur_map, inpath, outpath):
    
    klausur_map['PDFSeite_klausur'] = -1
    klausur_map['Filename_klausur'] = ''    
    
    try:
        os.makedirs(outpath)
    except OSError:
        pass
    

    for g, mg in klausur_map.groupby(level = 'Gruppe'):
        print('Gruppe', g)
        gpath = os.path.join(outpath, 'Gruppe{:02d}'.format(g))
        try:
            os.makedirs(gpath)
        except OSError:
            pass
        
        for k, mk in mg.groupby(level = 'UID'):
            pdflen = 0
            Nachname = '-'.join(mk.iloc[0].Nachname.split(' ')) + '_' + '-'.join(mk.iloc[0].Vorname.split(' '))
            outname = os.path.join(outpath, 'Gruppe{:02d}'.format(g), 'Klausur_{}_{:d}'.format(Nachname, k))
            try:
                os.makedirs(outname)
            except OSError:
                pass

            for i, r in mk.iterrows():
#                print(r)
                page = convert_from_path(os.path.join(inpath, r.Filename_aufgabe), 300,
                                          first_page = r.PDFSeite_aufgabe,
                                          last_page = r.PDFSeite_aufgabe)[0]
                
                pdflen +=1
                page.save(os.path.join(outname, 'page_{}.jpg').format(pdflen), 'JPEG')
                
    return klausur_map


######### Main Program
warnings.filterwarnings("ignore")

# Getrennte Unterordner für jede Gruppe erzeugen?
folder_per_gruppe = True 

# Pfad zu den korrigierten Klausuren, nach Aufgaben und Gruppe sortiert
path_aufgaben_korrigiert = 'klausuren_aufgaben_korrigiert'

# Pfad zu den korrigierten Klausuren, nach Gruppe und Klausur sortiert
path_klausuren_korrigiert = 'klausuren_einzeln_korrigiert'


klausur_map = pd.read_pickle('klausur_map_aufgabe.pkl')

# Sammle Informationen über Zusatzblätter
zusatz = pd.read_excel('zusatzblaetter.xlsx').dropna()
zusatz.columns = ['Filename', 'Pages']
zusatz.set_index('Filename', inplace = True)

idx = klausur_map.index.names
klausur_map.reset_index(inplace = True) 

for i, z in zusatz.iterrows():
    pl = str(z.Pages).split(',')
    print(i,pl)

    for p in pl:
        fname = klausur_map.Filename_aufgabe[(klausur_map.Aufgabe == 11) & 
                                             (klausur_map.PDFSeite_aufgabe == int(p))].values[0]       
        fname = os.path.join(os.path.split(fname)[0], i)
        klausur_map.Filename_aufgabe[(klausur_map.Aufgabe == 11) & 
                                     (klausur_map.PDFSeite_aufgabe == int(p))] = fname

klausur_map.set_index(idx, inplace = True)


#%% Sortiere korrigierte Klausuren: Für jede Gruppe ein Ordner, darin für jeden Teilnehemr eine Klausurdatei
print('**** Sortiere nach Teilnehmern ****')

tklausur_map = sort_by_gruppe_teilnehmer(klausur_map, path_aufgaben_korrigiert, 
                                         path_klausuren_korrigiert, 
                                         folder_per_gruppe = folder_per_gruppe)

print('**** Schreibe Datenbank ****')
tklausur_map.to_pickle('klausur_map_teilnehmer.pkl')

