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


# Define routes for PDF to Word and Word to PDF conversions
@app.route('/pdf/to-word', methods=['POST'])
def pdf_to_word_endpoint():
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename.endswith('.pdf'):
            input_pdf = file
            output_docx = 'output.docx'

            # Convert PDF to Word
            pdf_to_word(input_pdf, output_docx)

            return send_file(output_docx, as_attachment=True), 200
        else:
            return jsonify({'error': 'Invalid file format'}), 400

# Define route for Word to pdf conversion

@app.route('/word/to-pdf', methods=['POST'])
def word_to_pdf_endpoint():
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename.endswith('.docx'):
            input_docx = file
            output_pdf = 'output.pdf'

            # Convert Word to PDF
            word_to_pdf(input_docx, output_pdf)

            return send_file(output_pdf, as_attachment=True), 200
        else:
            return jsonify({'error': 'Invalid file format'}), 400               