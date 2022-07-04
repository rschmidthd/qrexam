# -*- coding: utf-8 -*-
"""
@author: ufriess

Adapted Robert Schmidt Jan/Feb 2022

Note: This assumes that the number of problem can be determined from
      the file name. If a problem is in the wrong pile, the klausur_map
      may get corrupted. *Can improve this in the future*!
      The number of the problem is determine in row 97.


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
    
    return pd.DataFrame(index = pd.MultiIndex(levels=[[],[],[]],
                             codes=[[],[],[]],
                             names=['num', 'Aufgabe', 'Seite']),
                        columns = ['PDFSeite', 'Filename', 'Rotation'])

def add_Klausur_map(klausur_map, num, Aufgabe, Seite, PDFSeite, Filename,
                    Rotation = 0,
                    verbose = False,
                    verify_integrity = True, inplace = True):
    	
    if inplace:
        result = klausur_map
    else:
        result = klausur_map.copy()
        
    assert all([e >= 0 for e in [num, Aufgabe, Seite, PDFSeite]])
    result.loc[num, Aufgabe, Seite] = {
        'PDFSeite': PDFSeite, 
        'Filename': os.path.basename(Filename),
        'Rotation': Rotation
        }
    
    if verbose:
        print(result.loc[num, Aufgabe, Seite])
    
    if not inplace:
        return result
    
    
def scan_Klausur_QR_map(filename, verbose = False):
    
    
    global images
    print ('start scan_Klausur_QR_map')
    
    klausur_map = create_Klausur_map()
        
    failed = pd.DataFrame(columns = ['PDFSeite', 'Filename'])
    
    files = glob.glob(filename)

    for f in files:
        fname = os.path.basename(f)

        aufgabe = int(fname[1:3])     # extract Aufgabe from file name
        print('Scanning QR codes in', fname)
        
        doc = fitz.open(f)
        
        # 1. this assumes that front and reverse sides are scanned
        #    and put next to each other in pdf (even number of pages!)
        # 2. it is also assumed that there is only a bar code on the
        #    front side
        # 3. the script will determine where the qr code is and
        #    treat the other side as reverse
        for page in range(0,len(doc)-1,2):
            n_err = 0
            error = True
            while error:
                try:
                    images1 = get_images_from_PDF_page(doc, page)
                    images2 = get_images_from_PDF_page(doc, page+1)
                    error = False
                except:
                    n_err += 1
                    if n_err > 3:
                        raise
            
            qr1 = []
            qr2 = []
            
            for image in images1:
                code = pyzbar.decode(image.convert(mode = '1', dither = 0), symbols=[ZBarSymbol.QRCODE])
                qr1 = qr1 + code
            for image in images2:
                code = pyzbar.decode(image.convert(mode = '1', dither = 0), symbols=[ZBarSymbol.QRCODE])
                qr2 = qr2 + code
                
            l1=len(qr1)
            l2=len(qr2)

            p1=0
            p2=0
            if (l1 == 1) or (l2 ==1):
                if l1==1:
                    data = qr1[0].data.decode("utf-8")
                    rotation = 0 if qr1[0].rect.top < 800 else 180
                    info = data.split(',')
                    p1=0+int(info[2]) # first scan is first page
                    p2=1+int(info[2]) # first scan is first page
                elif l2==1:
                    data = qr2[0].data.decode("utf-8")
                    rotation = 0 if qr2[0].rect.top < 800 else 180
                    info = data.split(',')
                    p1=1+int(info[2]) # second scan is first page
                    p2=0+int(info[2]) # second scan is first page
                else:
                    print ('problem, no front page found')
         
                # good diagnostic in output (-> pipe to output file)
                print (info,p1,p2)
                
                add_Klausur_map(klausur_map,
                    int(info[0]), # Nummer
                    int(aufgabe), # Aufgabe 
                    int(p1), # Vorder/Rückseite
                    page + 1,     # PDF Seite
                    fname,            # Filename 
                    Rotation = rotation,
                    inplace = True
                    )
                add_Klausur_map(klausur_map,
                    int(info[0]), # Nummer
                    int(aufgabe), # Aufgabe 
                    int(p2), # Vorder/Rückseite
                    page + 2,     # PDF Seite
                    fname,            # Filename 
                    Rotation = rotation,
                    inplace = True
                    )
 
                if verbose:
                    print(page, info)
            else:
                failed = failed.append({'PDFSeite': page + 1, 'Filename': fname}, ignore_index = True)
                if l1 == 0 and l2 == 0:
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
path_scanned = 'scans'

#%% Sammle Informationen von QR-Codes in Klausuren
print('ccanning QR codes')
klausur_map, failed = scan_Klausur_QR_map(path_scanned + '/*.pdf', verbose = False)

print('writing data base')
# create also an xls version for easy checking with libreoffice
klausur_map.to_pickle('klausur_map.pkl')
klausur_map.to_excel('klausur_map.xls')

# easier access for manual checking with xls file
#failed.to_pickle('klausur_qr_scan_failed.pkl')
failed.to_excel('klausur_qr_scan_failed.xls')

# Manuelle Zuweisung nicht erkannter Seiten
