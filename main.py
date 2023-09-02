from flask import Flask, request, send_file
from PyPDF2 import PdfReader
from docx import Document
import io

from werkzeug.datastructures import FileStorage

app = Flask(__name__)


@app.route('/', methods=['POST'])
def convert_pdf_to_docx():
    try:

        pdf_file: FileStorage = request.files['pdf']

        pdf_reader = PdfReader(pdf_file)

        doc = Document()

        for page in pdf_reader.pages:
            text = page.extract_text()
            doc.add_paragraph(text)

        doc_io = io.BytesIO()

        doc.save(doc_io)

        doc_io.seek(0)

        return send_file(doc_io, as_attachment=True)

    except KeyError:
        return "No file found in request", 400

    except Exception as e:
        return str(e), 500


if __name__ == '__main__':
    app.run(debug=True)
