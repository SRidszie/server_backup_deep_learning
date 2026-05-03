from django.shortcuts import render, redirect
from jd_parser.forms import ImageFileForm
from jd_parser.models import ImageFile
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
from .models import *
from django.views.generic import FormView
from rest_framework.exceptions import ParseError
from rest_framework.parsers import FileUploadParser
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.permissions import IsAuthenticated
from jd_parser import forms
from rest_framework.exceptions import APIException
from .utils import *
from rest_framework import status
from rest_framework_simplejwt.authentication import JWTAuthentication
from rest_framework.authentication import TokenAuthentication
import base64
from django.http.response import JsonResponse
from rest_framework.permissions import IsAuthenticated
from django.http import HttpResponse

# end

pytesseract.pytesseract.tesseract_cmd = (
    # "C:/Program Files/Tesseract-OCR/tesseract.exe"  # your path may be different
    "/usr/bin/tesseract"
)


class image_doc_docx_pdf_txt_rtf_to_json_view(APIView):
    permission_classes = (IsAuthenticated,)
    parser_class = (FileUploadParser,)

    def post(self, request):
        if "file_content" and "file_extension" not in request.data:
            # if "file_content" not in request.data:
            response = JsonResponse(
                {
                    "status": "failure",
                    "status-code": status.HTTP_400_BAD_REQUEST,
                    "message": "Request to add required parameters",
                }
            )
            return response

        file_content = request.data["file_content"]
        # file_extension = request.POST.get("file_extension")
        # file_extension = request.POST["file_extension"]

        file_extension = request.data["file_extension"]
        # print("Sandeep")
        # print(file_extension)
        # file_extension = request.POST.get("file_extension")

        text = process(file_content, file_extension)
        # text = extract_text_from_docx(file_content)
        return Response(text)


def valid_extensions_api(request):
    # data = {
    #     "Valid file extensions are: Word(.doc, .docx), Pdf(.pdf), Excel(.csv, .xlsx), Text(.txt, .rtf), Image(.jpg, .jpeg, .png)"
    # }

    data = {
        "valid_extensions": "doc, docx, pdf, csv, xlsx, txt, rtf, jpg, jpeg, png",
        "message": "Valid file extensions are: Word(.doc, .docx), Pdf(.pdf), Excel(.csv, .xlsx), Text(.txt, .rtf), Image(.jpg, .jpeg, .png)",
    }
    # data = {
    #     "PDF": ".pdf",
    #     "Word": [".doc", ".docx"],
    #     "Excel": [".csv", ".xlsx"],
    #     "Text": [".txt", ".rtf"],
    #     "Image": [".jpg", ".jpeg", ".png"],
    #     # "Others":".html"
    # }
    # return Response(data)
    return JsonResponse(data)


##################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ##########################################
#     # pdf api

# pdf with image api -->

# class pdf_with_image_to_view(APIView):
#     parser_class = [
#         MultiPartParser,
#     ]

#     def post(self, request, format=None):
#         if "name" not in request.data:
#             raise ParseError("Request to add required parameters")

#         filename = "temp_pdf_for_conversion.pdf"  # received file name
#         file_obj = request.data["name"]
#         with default_storage.open("jd_temp/" + filename, "wb+") as destination:
#             for chunk in file_obj.chunks():
#                 destination.write(chunk)
#         dirName = os.path.dirname(__file__)
#         cwd = Path.cwd()
#         filename = os.path.join(cwd, "media", "jd_temp/", "temp_pdf_for_conversion.pdf")
#         texts = extract_text_from_pdf_with_image(filename)
#         return Response(texts)


# Doc without image API -->
# class doc_without_image_to_view(APIView):
#     parser_class = [
#         MultiPartParser,
#     ]

#     def post(self, request, format=None):
#         if "name" not in request.data:
#             raise ParseError("Request to add required parameters")
#             return Response(status=status.HTTP_400_BAD_REQUEST)

#         filename = "temp_doc_for_conversion.doc"  # received file name
#         file_obj = request.data["name"]
#         with default_storage.open("jd_temp/" + filename, "wb+") as destination:
#             for chunk in file_obj.chunks():
#                 destination.write(chunk)
#         dirName = os.path.dirname(__file__)
#         cwd = Path.cwd()
#         filename = os.path.join(cwd, "media", "jd_temp/", "temp_doc_for_conversion.doc")
#         texts = extract_text_from_doc(filename)
#         return Response(texts, status=status.HTTP_200_OK)
