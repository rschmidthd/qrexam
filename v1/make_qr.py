# -*- coding: utf-8 -*-
"""
Created on Tue Jun 16 10:37:57 2020

@author: ufriess
"""

import qrcode
from PyPDF2 import PdfFileWriter, PdfFileReader

qr = qrcode.QRCode(
#     box_size=200,
#     border=4,
)
qr.add_data('Dies ist ein Test')
qr.make(fit=True)

qr = qr.make_image(fill_color="black", back_color="white")        

qr.save('qr.pdf', format = 'PDF', resolution = 600)


output = PdfFileWriter()
input1 = PdfFileReader(open("testpages\\testpage.pdf", "rb"))
qr = PdfFileReader(open("qr.pdf", "rb"))

output.addPage(input1.getPage(0))
output.getPage(0).mergePage(qr.getPage(0))

# finally, write "output" to document-output.pdf
with open("testpage_qr.pdf", "wb") as outputStream:
    output.write(outputStream)