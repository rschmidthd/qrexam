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
    xls file with participants from data base
    with
        - family name
        - given name
        - UID
        - matriculation number
        - exam number
"""

import pandas as pd
import os
from pdf2image import convert_from_path
from PyPDF2 import PdfFileWriter, PdfFileReader
import warnings


def sort_by_number(klausur_map, participants, inpath, outpath, folder_per_gruppe = False):
    
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

    # takes list of participants and extracts Name, ID and exam number
    #
    for k in range(len(participants)):   # len() gives number of entries
        mk=participants.iloc[k]
        pdflen = 0
        pdf_writer = PdfFileWriter()
        Nachname = '-'.join(mk.Name.split(' ')) + '_' + '-'.join(mk.Vorname.split(' '))
        num=mk.Nummer
        uid=mk.ID
        print (Nachname,num,uid)
            
        outname = 'Klausur_{}_{:d}.pdf'.format(Nachname, mk.ID)

        # once exam number is known, the solution pages can
        # be extracted from klausur_map for all problem numbers
        #
        for j in range(naufgaben+1):
            # sub-select all pages for given Aufgabe
            # the trick is pandas.IndexSlice[]
            #
            # OUCH! This crashes in reality test!!
                                                # tuple (num,aufgabe,seite)
            #sub_map = klausur_map.loc(axis=0)[pd.IndexSlice[num,j,:]]

            #this solution due to Clement Ranc works:
            sub_map = klausur_map.iloc[(klausur_map.index.get_level_values('num')==num)
                &(klausur_map.index.get_level_values('Aufgabe')==j)]

            pdflen += 1
            # goes through page numbers, which are the last of the
            # multiindex and thus sorted
            #print (j,sub_map.Filename_aufgabe)  # debugging
            for i, r in sub_map.iterrows():
                                              # tuple (num,aufgabe,seite)
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



######### Main Program
warnings.filterwarnings("ignore")

# path to input files (separate pdf for each problem)
path_problems = 'exams_marked'

# path final exams
path_exams = 'exams_out'

# Number of problems
naufgaben=8

results='results.xlsx'

participants = pd.read_excel(results, header = 3,
             usecols = ['ID', 'Matrikelnummer', 'Name', 'Vorname', 'Nummer' ])

#klausur_map = pd.read_pickle('github/new.pkl')
klausur_map = pd.read_pickle('klausur_map_aufgabe.pkl')


#%% sort marked exams, one output exam for each participants
print('**** sort by exam number ****')

tklausur_map = sort_by_number(klausur_map, participants,
                              path_problems, path_exams)

print('**** writing data base ****')
tklausur_map.to_pickle('klausur_map_teilnehmer.pkl')

