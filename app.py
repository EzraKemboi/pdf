from flask import Flask, request, jsonify, send_file
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import comtypes.client


app = Flask(__name__)

#Install required dependencies
app.config['MAX_CONTENT_LENGHT']= 16 * 1024 #set maximum file size to 16MB
# Write file manipulation functions
def pdf_to_word(input_pdf, output_pdf):
    #Read the pdf file
    with open(input_pdf, 'rb') as f:
        pdf_reader = PyPDF2.Pdfreader(f)

        #create a new word document
        document = docx.Document()

        #Extract text from each page in the PDF
        for page in pdf_reader.pages:
            text= page.extract_text()
            document.add_paragraph(text)

        #Save the Word document
         document.save(output_docx)

    def word_to_pdf(input_docx, out_pdf):
        #Create a COM object for Word 
        word = comtypes.client.CreateObject('Word.Application')
        # Open the word document
        doc =word.Documents.Open(input_docx)
        
        #Convert the document to PDF
        doc.SaveAs(output_pdf, FileFormat = 17)

        #Close the document and quit word
        doc.Close()
        word.Quit()


               