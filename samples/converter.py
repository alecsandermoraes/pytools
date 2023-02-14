from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter 
from pdfminer.converter import TextConverter 
from pdfminer.layout import LAParams 
from pdfminer.pdfpage import PDFPage 

from reportlab.pdfgen.canvas import Canvas 
from pptx import Presentation 
from docx import Document 
from io import StringIO 
from io import BytesIO 
from PIL import Image 
import pandas as pd 
import pypandoc 
import sys, os 

class Image():
    
    class JIF():
        
        def __init__(self, filename):
            self.filename = filename 
            
        def ToJFIF(self, outputfile):
            with Image.open(self.filename) as img:
                img.save(outputfile, 'JPEG', quality = 95, jfif_version = (1, 1))
            return True 
        
        def ToJPG(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGB')
                img.save(outputfile, 'JPEG', quality = 95)
            return True 
        
        def ToPNG(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'PNG')
            return True 
        
        def ToPSD(self, outputfile):
            with Image.open(self.filename) as img:
                img.save(outputfile, 'PSD')
            return True 
        
        def ToTIFF(self, outputfile):
            with Image.open(self.filename) as img:
                img.save(outputfile, 'TIFF')
            return True 
        
        def ToWEBP(self, outputfile):
            with Image.open(self.filename) as img:
                img.save(outputfile, 'WEBP', quality = 95)
            return True 
    
    class JFIF():
        
        def __init__(self, filename):
            self.filename = filename 
            
        def ToJIF(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGB')
                img.save(outputfile, 'JIF')
            return True 
        
        def ToJPG(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGB')
                img.save(outputfile, 'JPEG')
            return True 
        
        def ToPNG(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'PNG')
            return True 
        
        def ToPSD(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'PSD')
            return True 

        def ToSVG(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'SVG')
            return True 
        
        def ToTIFF(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'TIFF')
            return True 
        
        def ToWEBP(self, outputfile):
            with Image.open(self.filename) as img:
                img = img.convert('RGBA')
                img.save(outputfile, 'WEBP')
            return True 

# PPTX & PDF 
def PPTXToPDF(inputfile, outputfile):
    prs = Presentation(inputfile)
    prs.save(outputfile, 'pdf')
    return True 

def PDFToPPTX(inputfile, outputfile):
    manager = PDFResourceManager() 
    textio = StringIO()
    laparams = LAParams() 
    converter = TextConverter(manager, textio, laparams = laparams)
    interpreter = PDFPageInterpreter(manager, converter)
    
    with open(inputfile, 'rb') as file:
        for page in PDFPage.get_pages(file):
            interpreter.process_page(page)
    text = textio.getvalue() 
    
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    text_box = slide.shapes.add_textbox(left = 0, top = 0, width = prs.slide_width, height = prs.slide_height)
    text_frame = text_box.text_frame 
    text_frame.clear() 
    text_frame.text = text 
    
    prs.save(outputfile)
    return True 

# XLSX & CSV 
def XLSXToCSV(inputfile, outputfile):
    df = pd.read_excel(inputfile)
    df.to_csv(outputfile, index = False)
    return True 

def CSVToXLSX(inputfile, outputfile):
    df = pd.read_csv(inputfile)
    df.to_excel(outputfile, index = False)
    return True 

# DOCX & PDF 
def DOCXToPDF(inputfile, outputfile):
    doc = Document(inputfile)
    pdffile = BytesIO()
    canvas_obj = Canvas(pdffile)
    
    for para in doc.paragraphs:
        canvas_obj.drawString(0, 0, para.text)
    canvas_obj.save()
    
    with open(outputfile, 'wb') as file:
        file.write(pdffile.getvalue())
        file.close()
    return True 

def PDFToDOCX(inputfile, outputfile):
    pypandoc.convert_file(inputfile,
    'docx', outputfile = outputfile)
    return True