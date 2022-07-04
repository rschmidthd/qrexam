# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 09:41:35 2020

@author: ufriess

Required packages:
    pytzbar   - pip install pyzbar
    fitz      - pip install PyMuPDF
    PIL       - pip install pillow
    pdf2image - pip install pdf2image
"""

import pandas as pd
import os
import io
from pyzbar import pyzbar
from pyzbar.pyzbar import ZBarSymbol
import fitz
from PIL import Image
import glob
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



######### Main Program
warnings.filterwarnings("ignore")


# Pfad zu den gescannten Klausuren 
#path_scanned = 'klausuren_druckvorlage' # Testmodus
path_scanned = 'klausuren_scanned'


#%% Sammle Informationen von QR-Codes in Klausuren
print('Scanning QR codes')
klausur_map, failed = scan_Klausur_QR_map(path_scanned + '\\*.pdf', verbose = False)

print('Schreibe Datenbank')
klausur_map.to_pickle('klausur_map.pkl')
#failed.to_pickle('klausur_qr_scan_failed.pkl')
failed.to_excel('klausur_qr_scan_failed.xls')

# Manuelle Zuweisung nicht erkannter Seiten
