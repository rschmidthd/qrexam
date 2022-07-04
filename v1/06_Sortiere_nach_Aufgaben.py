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
import os
import io
from pyzbar import pyzbar
from pyzbar.pyzbar import ZBarSymbol
import fitz
from PIL import Image
import glob
import qrcode
from pdf2image import convert_from_path
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import warnings

def get_images_from_PDF_page(doc, page):
    
    result = []
    for i, img in enumerate(doc.getPageImageList(page)):
        xref = img[0]
        pix = fitz.Pixmap(doc, xref)
        if pix.n < 5:       # this is GRAY or RGB
            d = pix.getPNGdata()
        else:               # CMYK: convert to RGB first
            pix1 = fitz.Pixmap(fitz.csRGB, pix)
            d = pix1.getPNGdata()
            
        fimg = io.BytesIO()
        fimg.write(d)
        fimg.seek(0)
        result = result + [Image.open(fimg)]
        
#        os.remove('tmp.png')
        
    return result

def create_Klausur_map():
    
    return pd.DataFrame(index = pd.MultiIndex(levels=[[],[],[],[]],
                             codes=[[],[],[],[]],
                             names=['UID', 'Aufgabe', 'Seite', 'Gruppe']),
                        columns = ['PDFSeite', 'Filename', 'Rotation'])

def add_Klausur_map(klausur_map, UID, Aufgabe, Seite, Gruppe, PDFSeite, Filename,
                    Rotation = 0,
                    verbose = False,
                    verify_integrity = True, inplace = True):
    	
    if inplace:
        result = klausur_map
    else:
        result = klausur_map.copy()
        
    assert all([e >= 0 for e in [UID, Aufgabe, Seite, Gruppe, PDFSeite]])
    result.loc[UID, Aufgabe, Seite, Gruppe] = {
        'PDFSeite': PDFSeite, 
        'Filename': os.path.basename(Filename),
        'Rotation': Rotation
        }
    
    if verbose:
        print(result.loc[UID, Aufgabe, Seite, Gruppe])
    
    if not inplace:
        return result
    
    
def scan_Klausur_QR_map(filename, verbose = False):
    
    
    global images
    
    
    klausur_map = create_Klausur_map()
        
    failed = pd.DataFrame(columns = ['PDFSeite', 'Filename'])
    
    files = glob.glob(filename)

    for f in files:
        fname = os.path.basename(f)
        print('Scanning QR codes in', fname)
        
        doc = fitz.open(f)
        
        for page in range(len(doc)):
            n_err = 0
            error = True
            while error:
                try:
                    images = get_images_from_PDF_page(doc, page)
                    error = False
                except:
                    n_err += 1
                    if n_err > 3:
                        raise
            
            qr = []
            
            for image in images:
                qr = qr + pyzbar.decode(image.convert(mode = '1', dither = 0), symbols=[ZBarSymbol.QRCODE])
                
            if len(qr) == 1:
                data = qr[0].data.decode("utf-8")
                info = data.split(',')
                
                rotation = 0 if qr[0].rect.top < 800 else 180

                add_Klausur_map(klausur_map,
                    int(info[0]), # UID 
                    int(info[2]), # Aufgabe 
                    int(info[3]), # Seite
                    int(info[1]), # Gruppe
                    page + 1,     # PDF Seite
                    fname,            # Filename 
                    Rotation = rotation,
                    inplace = True
                    )
 
                if verbose:
                    print(page, info)
            else:
                failed = failed.append({'PDFSeite': page + 1, 'Filename': fname}, ignore_index = True)
                if len(qr) == 0:
                    print('No QR code found on page',page + 1)
                else:
                    print('Multiple QR codes found on page',page + 1)
               
            # for img in images:
            #     img.close()
                    
        doc.close()
    
    return klausur_map.drop_duplicates().sort_index(), failed
 
def add_missing_qr(missing_list, inpath, outpath):
    """
    Add missing QR codes on specific pages of a PDF file

    Parameters
    ----------
    missing_list : pandas.DataFrame
        List of missing data including page number and qr code data.
        index:  UID, Aufgabe, Seite, Gruppe
        columns: PDFSeite, Filename.
    suffix : str
        Suffix to be added to the output filenames
        
    Returns
    -------
    None.

    """
    
    try:
        os.makedirs(outpath)
    except OSError:
        pass
        
    for f, g in missing_list.groupby('Filename'):
        infile = os.path.join(inpath, f)    
        outfile = os.path.join(outpath, f)   
        
        f_out = PdfFileWriter()
        f_in = PdfFileReader(open(infile, 'rb'))
    
     
        # copy all pages to output
        for page in f_in.pages:       
            f_out.addPage(page)
    
    #     for i, r in missing_list.iterrows():
    # #\qrcode{\UIDnummer,\gruppe,\arabic{section},\the\numexpr\value{page}+1\relax}    
    #         qr = qrcode.QRCode(
    #         #     box_size=200,
    #         #     border=4,
    #         )
    #         #                    UID    Gruppe     Aufgabe    Seite                           
    #         qr.add_data(','.join([str(i[0]), str(i[3]), str(i[1]), str(i[2])]))
    #         qr.make(fit=True)
            
    #         qr = qr.make_image(fill_color="black", back_color="white")   
    #         width, height = qr.size
    #         img = Image.new('RGB', (int(1.5*width), height), (255, 255, 255))
    #         img.paste(qr)
    #         img.save('qr_tmp.pdf', format = 'PDF', resolution = 100)
            
    #         fqr = PdfFileReader(open('qr_tmp.pdf', 'rb'))
        
    #         f_out.getPage(r.PDFSeite - 1).mergePage(fqr.getPage(0))
        
        for i, r in missing_list.iterrows():
            qrpage = io.BytesIO()
        
            can = canvas.Canvas(qrpage, pagesize=A4)
        
            qr = qrcode.QRCode(
             #     box_size=200,
             #     border=4,
             )
            #                    UID    Gruppe     Aufgabe    Seite                           
            qr.add_data(','.join([str(i[0]), str(i[3]), str(i[1]), str(i[2])]))
            qr.make(fit=True)
            qr = qr.make_image(fill_color="black", back_color="white")   
            
            io_img = io.BytesIO()
            qr.save(io_img, 'PNG')
            
            reportlab_io_img = ImageReader(io_img)
            
            can.drawImage(reportlab_io_img, 30, 25, width = 60, 
                          preserveAspectRatio = True,
                          anchor = 'sw')
    
            can.save()
            
            # move to the beginning of the StringIO buffer
            qrpage.seek(0)
            qr_pdf = PdfFileReader(qrpage)
            
            # add QR code
            # page.mergePage(qr_pdf.getPage(0))
            # output.addPage(page)
            qr_pdf.getPage(0).mergePage(page)
            f_out.getPage(r.PDFSeite - 1).mergePage(qr_pdf.getPage(0))
        
        # finally, write "output" to document-output.pdf
        with open(outfile, 'wb') as outputStream:
            f_out.write(outputStream)    



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


def sort_by_aufgabe_gruppe(klausur_map, inpath, outpath, folder_per_gruppe = True):
    
    global i,r
     
    klausur_map['PDFSeite_aufgabe'] = -1
    klausur_map['Filename_aufgabe'] = ''    
    
    try:
        os.makedirs(outpath)
    except OSError:
        pass
    
    # Create PDFs in dictionary
    PDFs = {}
    for f, g in klausur_map.groupby('Filename'):
        PDFs[f] = PdfFileReader(os.path.join(inpath, f))

    for a, ma in klausur_map.groupby(level = 'Aufgabe'):
        print('Aufgabe', a)
        ma = ma.sort_values(by = ['Nachname', 'Vorname'])
        
        apath = 'Aufgabe{:02d}'.format(a)
        
        
        if folder_per_gruppe: # für jede Gruppe einen Unterordner
            try:
                os.makedirs(os.path.join(outpath, apath))
            except OSError:
                pass

            for g, mg in ma.groupby(level = 'Gruppe'):
                pdflen = 0
                pdf_writer = PdfFileWriter()
                outname = os.path.join(apath, 'Klausur_A{:02d}_G{:02d}.pdf'.format(a, g))
    
                for i, r in mg.iterrows():
    #                print(r)
                    pdf_writer.addPage(PDFs[r.Filename].getPage(r.PDFSeite - 1).rotateCounterClockwise(r.Rotation)) 
                    
                    pdflen += 1
                    klausur_map.PDFSeite_aufgabe.loc[i] = pdflen
                    klausur_map.Filename_aufgabe.loc[i] = outname
                    
                with open(os.path.join(outpath, outname), 'wb') as out:
                    pdf_writer.write(out)
        else: # alle Blätter einer Aufgabe in eine Datei
            pdflen = 0
            pdf_writer = PdfFileWriter()
            outname = 'Klausur_A{:02d}.pdf'.format(a)

            for i, r in ma.iterrows():
#                print(r)
                pdf_writer.addPage(PDFs[r.Filename].getPage(r.PDFSeite - 1).rotateCounterClockwise(r.Rotation)) 
                
                pdflen += 1
                klausur_map.PDFSeite_aufgabe.loc[i] = pdflen
                klausur_map.Filename_aufgabe.loc[i] = outname
                
            with open(os.path.join(outpath, outname), 'wb') as out:
                pdf_writer.write(out)
                
    return klausur_map


def sort_by_gruppe_teilnehmer(klausur_map, inpath, outpath):
    
    klausur_map['PDFSeite_klausur'] = -1
    klausur_map['Filename_klausur'] = ''    
    
    try:
        os.makedirs(outpath)
    except OSError:
        pass
    
    # Create PDFs in dictionary
    PDFs = {}
    for f, g in klausur_map.groupby('Filename_aufgabe'):
        PDFs[f] = PdfFileReader(os.path.join(inpath, f))


    for g, mg in klausur_map.groupby(level = 'Gruppe'):
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
            outname = os.path.join('Gruppe{:02d}'.format(g), 'Klausur_{}_{:d}.pdf'.format(Nachname, k))

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

def check_corrected_exams(klausur_map, path):

    missing = []

    for a, ma in klausur_map.groupby(level = 'Aufgabe'):
        apath = os.path.join(path, 'Aufgabe{:02d}'.format(a))
        
        for g, mg in ma.groupby(level = 'Gruppe'):
            fname = os.path.join(apath, 'Klausur_A{:02d}_G{:02d}.pdf'.format(a, g))
            if not os.path.exists(fname):
                print(os.path.basename(fname), 'nicht vorhanden')
                missing = missing + [fname]

    return(missing)    


def rotate_pdf(infile, outfile, angle = 180, skip_pages = []):
    
    pdf_in = open(infile, 'rb')
    pdf_reader = PdfFileReader(pdf_in)
    pdf_writer = PdfFileWriter()
    for pagenum in range(pdf_reader.numPages):
        page = pdf_reader.getPage(pagenum)
        if not (pagenum + 1 in skip_pages):
            page.rotateClockwise(angle)
        pdf_writer.addPage(page)
    pdf_out = open(outfile, 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()



######### Main Program
warnings.filterwarnings("ignore")

# Getrennte Unterordner für jede Gruppe erzeugen?
folder_per_gruppe = True

# Pfad zu den gescannten Klausuren 
#path_scanned = 'klausuren_druckvorlage' # Testmodus
path_scanned = 'klausuren_scanned'

# Pfad zu den Klausuren nach Aufgabe sortiert
path_aufgaben = 'klausuren_aufgaben_unkorrigiert'

# Pfad zu den korrigierten Klausuren, nach Aufgaben und Gruppe sortiert
path_aufgaben_korrigiert = 'klausuren_aufgaben_korrigiert'

# Pfad zu den korrigierten Klausuren, nach Gruppe und Klausur sortiert
path_klausuren_korrigiert = 'klausuren_einzeln_korrigiert'

num_aufgaben = 12
seiten_pro_aufgabe = [2 if a < 11 else 6 for a in range(num_aufgaben)]

#%% Sortiere Klausuren: Für jede Aufgabe ein Ordner, darin für jede Aufgabe eine Datei
#   --> Für Korrekturen
klausur_map = pd.read_pickle('klausur_map_teilnehmer.pkl')

print('**** Sortiere nach Aufgaben ****')
klausur_map = sort_by_aufgabe_gruppe(klausur_map, path_scanned, path_aufgaben, 
                                     folder_per_gruppe = folder_per_gruppe)

print('Schreibe Datenbank')
klausur_map.to_pickle('klausur_map_aufgabe.pkl')

#%% Erzeuge Ordner für korrigierte Aufgaben

if folder_per_gruppe:
    for a in range(num_aufgaben):
        apath = 'Aufgabe{:02d}'.format(a)
        os.makedirs(os.path.join(path_aufgaben_korrigiert, apath), exist_ok = True)
else:
    os.makedirs(path_aufgaben_korrigiert, exist_ok = True)
    

