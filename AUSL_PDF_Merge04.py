
#  extracting_text.py estrarre testo dalla prima pagina di un pdf
# testato con V 2.7
#
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter

from reportlab import *
from reportlab.lib.utils import *
from reportlab.lib.colors import CMYKColor, PCMYKColor
from reportlab.pdfgen.canvas import Canvas

import sys
import os

import locale
locale.setlocale(locale.LC_ALL, 'ita_ita')
print locale.getlocale(locale.LC_TIME)


############ Formato A4: 595 x 842 ############

############ Creo una pagina vuota ############
emptypage = "emptypage.pdf"
c = Canvas(emptypage)
c.showPage()
c.save()

# walk_dir = sys.argv[1]
# walk_dir = "d:\\_AUSL\\StampeCollegio201911\\CCP 2\\"
walk_dir = "D:\\_AUSL\\StampeCollegio201911\\ZimbraItems\\Inbox\GIORNALI DI CASSA AUSL ROMAGNA 3 TRIM 2019\\"
#print('walk_dir = ' + walk_dir)

# If your current working directory may change during script execution, it's recommended to
# immediately convert program arguments to an absolute path. Then the variable root below will
# be an absolute path as well. Example:
# walk_dir = os.path.abspath(walk_dir)
print('walk_dir (absolute) = ' + os.path.abspath(walk_dir))

if not os.path.isdir(walk_dir):
    print "Attenzione la directory non esiste!!!"

merger = PdfFileMerger()
banner = PdfFileWriter()
contatore = 0

for root, subdirs, files in os.walk(walk_dir):
    for subdir in subdirs:
        #print('\t- subdirectory ' + subdir)
        pass
    for filename in files:
        file_path = os.path.join(root, filename)
        #print('\t- file %s (full path: %s)' % (filename, file_path))
        #base = filename.split( '.')
        #if base[1]=="pdf":
        #    percorso = file_path
        #    text_extractor(percorso)
        if filename.endswith(".pdf") or filename.endswith(".PDF"):
            print "Processo il file " + file_path

            # valuto il numero di pagine
            aggiungipagina = 0
            inputpdf = open(file_path, 'rb')
            pdf = PdfFileReader(inputpdf)
            numerodipagine = pdf.getNumPages()
            inputpdf.close()
            print "Pagine: " + str(numerodipagine)
            if (numerodipagine % 2) != 0:
                print "occorre aggiungere una pagina!"
                aggiungipagina = 1

            #creo e inserisco i dati nel file banner con reportlab
            bannerfile = "temp_pdf_" + str(contatore) + ".pdf"
            c = Canvas(bannerfile)
            #le info sotto vanno perse con il merger successivo....
            #c.setAuthor("Pierangelo Motta")
            #c.setTitle("Merge Pdf e aggiunta Banner")

            #queste le ho copiate da Tarros, ma non servono qui...
            #import reportlab.rl_config
            #reportlab.rl_config.warnOnMissingFontGlyphs = 0
            #from reportlab.pdfbase import pdfmetrics
            #from reportlab.pdfbase.ttfonts import TTFont
            #pdfmetrics.registerFont(TTFont('LucidaConsole', 'lucon.ttf'))

            blue = PCMYKColor(100, 91, 30, 22)
            c.setStrokeColor(blue)
            c.setLineWidth(2)
            c.line(30, 743, 565, 743)
            c.setFillColor(blue)
            c.setFontSize(14)

            #c.drawString(258, 716, str(filename))
            c.drawString(30, 716, "Nome File: ")
            c.drawString(120, 716, unicode(filename, '1252'))

            stringadata = str(time.strftime("%A %d/%m/%Y - %H:%M:%S", time.localtime()))
            stringadataUni = unicode(stringadata, '1252')

            c.drawString(30, 691, "Data: ")
            c.drawString(120, 691, stringadataUni.capitalize())

            c.drawString(30, 666, "N. pag: ")
            c.drawString(120, 666, str(numerodipagine))

            ##   Here, % Y, % m, % d, % H  etc.are format codes.
            #   % Y - year[0001, ..., 2018, 2019, ..., 9999]
            #   % m - month[01, 02, ..., 11, 12]
            #   % d - day[01, 02, ..., 30, 31]
            #   % H - hour[00, 01, ..., 22, 23
            #   % M - month[00, 01, ..., 58, 59]
            #   % S - second[00, 01, ..., 58, 61]

            c.showPage()
            c.save()


            bannerfile_per_merger = open(bannerfile, "rb")
            merger.append(bannerfile_per_merger)
            merger.append(emptypage)
            bannerfile_per_merger.close()

            inputpdf = open(file_path, "rb")
            if numerodipagine < 90:
                merger.append(inputpdf)
            else:
                merger.append(inputpdf, pages=(1,))
                merger.append(inputpdf, pages=(numerodipagine -1, numerodipagine ))

            if aggiungipagina == 1 and numerodipagine < 90:
                merger.append(emptypage)

            inputpdf.close()

            os.remove(bannerfile)
            contatore = contatore + 1

# Write to an output PDF document
output = open("document-output.pdf", "wb")
infos = {'/Title': 'Merger Pdf con Banner'}
merger.addMetadata(infos)
infos = {'/Author': 'Pierangelo Motta'}
merger.addMetadata(infos)



merger.write(output)

merger.close()

os.remove(emptypage)
