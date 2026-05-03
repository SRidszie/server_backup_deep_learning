from django.shortcuts import render, redirect
from resume_parser.forms import ImageFileForm
from resume_parser.models import ImageFile
from .forms import ImageFileForm
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
import base64
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
from pathlib import Path
from rest_framework.viewsets import ViewSet
from rest_framework.response import Response
from rest_framework import generics
from rest_framework.views import APIView

from rest_framework.views import APIView
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework.response import Response
from rest_framework import status

# from .serializers import FileSerializer
from .models import *


from django.views.generic import FormView
from rest_framework.exceptions import ParseError
from rest_framework.parsers import FileUploadParser
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.permissions import IsAuthenticated
from resume_parser import forms

from .utils import *
from rest_framework import status
from rest_framework_simplejwt.authentication import JWTAuthentication

# end
pytesseract.pytesseract.tesseract_cmd = (
    # "C:/Program Files/Tesseract-OCR/tesseract.exe"  # your path may be different
    "/usr/bin/tesseract"
)


def upload_doc(request):
    # # create object of form
    form = UploadDocumentForm(request.POST or None, request.FILES or None)

    # check if form data is valid
    if form.is_valid():
        form.save()

    context = UploadDocumentForm()

    return render(request, "jd_parser/upload_doc.html", {"form": context})


def home(request):
    data = dict()

    image_form = ImageFileForm(request.POST or None, request.FILES or None)
    if image_form.is_valid():
        image = image_form.save()
        image.execute_and_save_ocr()

        redirect("home")

    image_list = ImageFile.objects.all().order_by("-id")

    data["image_form"] = image_form
    data["image_list"] = image_list
    return render(request, "jd_parser/index.html", data)


# image API
# local system image
class OcrDataApiView(APIView):
    parser_class = (FileUploadParser,)

    def post(self, request):
        if "name" not in request.data:
            raise ParseError("Request to add required parameters")

        f = request.data["name"]

        text = extract_text_from_image(f)
        return Response(text)

    # pdf api


class pdf_with_image_to_view(APIView):
    parser_class = [
        MultiPartParser,
    ]

    def post(self, request, format=None):
        if "name" not in request.data:
            raise ParseError("Request to add required parameters")

        filename = "temp_pdf_for_conversion.pdf"  # received file name
        file_obj = request.data["name"]
        with default_storage.open("jd_temp/" + filename, "wb+") as destination:
            for chunk in file_obj.chunks():
                destination.write(chunk)
        dirName = os.path.dirname(__file__)
        cwd = Path.cwd()
        filename = os.path.join(cwd, "media", "jd_temp/", "temp_pdf_for_conversion.pdf")
        texts = extract_text_from_pdf_with_image(filename)
        return Response(texts)


# Api call for multipurposse which accept docx,pdf,image with all extensions
class img_pdf_docx_to_json_view(APIView):
    permission_classes = (IsAuthenticated,)
    parser_class = FileUploadParser

    def post(self, request, format=None):
        if "name" not in request.data:
            raise ParseError("Request to add required parameters")
            return Response(status=status.HTTP_400_BAD_REQUEST)

        f = request.data["name"]
        text = process(f)

        return Response(text, status=status.HTTP_200_OK)


# Doc API
class doc_without_image_to_view(APIView):
    parser_class = [
        MultiPartParser,
    ]

    def post(self, request, format=None):
        if "name" not in request.data:
            raise ParseError("Request to add required parameters")
            return Response(status=status.HTTP_400_BAD_REQUEST)

        filename = "temp_doc_for_conversion.doc"  # received file name
        file_obj = request.data["name"]
        with default_storage.open("jd_temp/" + filename, "wb+") as destination:
            for chunk in file_obj.chunks():
                destination.write(chunk)
        dirName = os.path.dirname(__file__)
        cwd = Path.cwd()
        filename = os.path.join(cwd, "media", "jd_temp/", "temp_doc_for_conversion.doc")
        texts = extract_text_from_doc(filename)
        return Response(texts, status=status.HTTP_200_OK)
