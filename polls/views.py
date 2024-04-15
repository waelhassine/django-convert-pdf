from django.http import JsonResponse, FileResponse
from spire.doc import Document
from spire.xls import Workbook
from spire.presentation import Presentation
import tempfile
import os

def convert_to_pdf(request):
    if request.method == 'POST' and request.FILES.get('file'):
        uploaded_file = request.FILES['file']
        file_extension = uploaded_file.name.split('.')[-1].lower()

        # Check file extension and convert to PDF accordingly
        if file_extension in ['doc', 'docx']:
            document = Document()
            document.load_from_stream(uploaded_file)
            pdf_path = convert_doc_to_pdf(document)

        elif file_extension in ['xls', 'xlsx']:
            workbook = Workbook()
            workbook.load_from_stream(uploaded_file)
            pdf_path = convert_excel_to_pdf(workbook)

        elif file_extension in ['ppt', 'pptx']:
            presentation = Presentation()
            presentation.load_from_stream(uploaded_file)
            pdf_path = convert_ppt_to_pdf(presentation)

        else:
            return JsonResponse({'error': 'Unsupported file format'})

        # Check if PDF conversion was successful
        if pdf_path:
            # Return the PDF file as a response
            with open(pdf_path, 'rb') as f:
                response = FileResponse(f, content_type='application/pdf')
                response['Content-Disposition'] = 'attachment; filename=output.pdf'
            return response
        else:
            return JsonResponse({'error': 'Failed to convert file to PDF'})

    else:
        return JsonResponse({'error': 'No file uploaded or invalid request method'})

def convert_doc_to_pdf(document):
    pdf_path = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False).name
    document.save_to_file(pdf_path)
    return pdf_path

def convert_excel_to_pdf(workbook):
    pdf_path = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False).name
    workbook.save_to_file(pdf_path)
    return pdf_path

def convert_ppt_to_pdf(presentation):
    pdf_path = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False).name
    presentation.save_to_file(pdf_path)
    return pdf_path
