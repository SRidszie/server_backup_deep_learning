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

# from .serializers import FileSerializer
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
import base64
from django.http.response import JsonResponse

# from .exceptions import JDParserResponse

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
# from django.http import HttpResponse
# import responses


# class HttpResponseNoContent(HttpResponse):
#     status_code = 200


class image_doc_docx_pdf_txt_rtf_to_json_view(APIView):
    parser_class = (FileUploadParser,)

    def post(self, request):
        if "file_content" and "file_type" not in request.data:
            response = JsonResponse(
                {
                    "status": "failure",
                    "status-code": status.HTTP_400_BAD_REQUEST,
                    "message": "Request to add required parameters",
                }
            )
            return response
            # raise ParseError("Request to add required parameters")
        # responses.status_code = 200
        # if responses is not None:
        #     responses.status["status_code"] = responses.status_code
        # if "file_content" is not None:
        #     raise JDParserResponse()
        file_type = request.POST.get("file_type")
        # print(file_type)
        file_content = request.data["file_content"]

        text = extract_content_filetype(file_content, file_type)
        return Response(text)
        # sample_string_bytes = f.encode("ascii")
        # # sample_string = f.decode("utf-8")
        # base64_encoded_data = base64.b64encode(sample_string_bytes)
        # base64_message = base64_encoded_data.decode("utf-8")
        # # base64_message = base64_encoded_data.decode("ascii")
        # # # base64_bytes = base64.b64encode(sample_string_bytes)
        # # # base64_string = base64_bytes.decode("ascii")
        # base64_bytes_ = base64.b64decode(base64_message)
        # base64_string_ = base64_bytes_.decode("utf-8")
        # print(base64_string_)
        # print(type(f))

        # sample_string_bytes = base64.b64decode(f)
        # sample_string = sample_string_bytes.decode("utf-8")
        # print(sample_string)

        # base64_string = f.decode("ascii")
        # base64_bytes_ = base64.b64decode(base64_string)
        # base64_string_ = base64_bytes_.decode("ascii")
        # print(base64_bytes_)

        # name = file_content(base64)
        # file_ext = (.pdf/.doc/.docx/.png/.jpg/.jpeg/.txt/.rtf)
        # return Response(serializer.data, status=status.HTTP_201_CREATED)
        # return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        # return HttpResponseNoContent()

        # text = extract_text_from_image(file_content)
        # text = extract_text_from_docx(file_content)
        # Process Fn:
        # text = process(file_content)
        # return Response(text)
        # For Image Fn:
        # if file_type == ".png" or ".PNG" or ".jpg" or ".JPG" or ".jpeg" or ".JPEG":
        #     text = extract_text_from_image(file_content)
        #     return Response(text)
        # if file_type == ".docx":
        #     text = extract_text_from_docx(file_content)
        #     return Response(text)
        # # elif file_type == ".pdf":
        #     text = extract_text_from_image(file_content)
        #     return Response(text)
        # elif file_type == ".doc":
        #     text = extract_text_from_image(file_content)
        #     return Response(text)

    # def post(self, request, format=None):
    #     if "name" not in request.data:
    #         raise ParseError("Request to add required parameters")

    #     filename = "temp_pdf_for_conversion.pdf"  # received file name
    #     file_obj = request.data["name"]
    #     with default_storage.open("jd_temp/" + filename, "wb+") as destination:
    #         for chunk in file_obj.chunks():
    #             destination.write(chunk)
    #     dirName = os.path.dirname(__file__)
    #     cwd = Path.cwd()
    #     filename = os.path.join(cwd, "media", "jd_temp/", "temp_pdf_for_conversion.pdf")
    #     texts = extract_text_from_pdf_with_image(filename)
    #     return Response(texts, status=status.HTTP_200_OK)

    # def post(self, request, format=None):
    #     if "name" not in request.data:
    #         raise ParseError("Request to add required parameters")
    #         return Response(status=status.HTTP_400_BAD_REQUEST)

    #     filename = "temp_doc_for_conversion.doc"  # received file name
    #     file_obj = request.data["name"]
    #     with default_storage.open("jd_temp/" + filename, "wb+") as destination:
    #         for chunk in file_obj.chunks():
    #             destination.write(chunk)
    #     dirName = os.path.dirname(__file__)
    #     cwd = Path.cwd()
    #     filename = os.path.join(cwd, "media", "jd_temp/", "temp_doc_for_conversion.doc")
    #     texts = extract_text_from_doc(filename)
    #     return Response(texts, status=status.HTTP_200_OK)


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


# # Api call for multipurposse which accept docx,pdf,image with all extensions
# class img_pdf_docx_to_json_view(APIView):
#     permission_classes = (IsAuthenticated,)
#     parser_class = FileUploadParser

#     def post(self, request, format=None):
#         if "name" not in request.data:
#             raise ParseError("Request to add required parameters")
#             return Response(status=status.HTTP_400_BAD_REQUEST)

#         f = request.data["name"]
#         text = process(f)

#         return Response(text, status=status.HTTP_200_OK)


# Api call for multipurpose which accept pdf,pdf+image, image, doc, docx with all extensions  -->


class img_pdf_docx_to_json_view(APIView):
    permission_classes = (IsAuthenticated,)
    # parser_class = FileUploadParser
    parser_class = [MultiPartParser, FileUploadParser]

    def post(self, request, format=None):
        if "name" not in request.data:
            raise ParseError("Request to add required parameters")

        f = request.data["name"]
        text = process(f)
        abc = f.str()
        # sample_string = "My Name is Sandeep"
        sample_string_bytes = abc.encode("ascii")  # bytes like character
        base64_bytes = base64.b64encode(sample_string_bytes)
        base64_string = base64_bytes.decode("ascii")
        base64_bytes_ = base64.b64decode(base64_string)
        base64_string_ = base64_bytes_.decode("ascii")

        print(f"Encoded string: {base64_string_}")

        # print("Sandeep")
        return Response(text, status=status.HTTP_200_OK)

    # def post(self, request, format=None):
    #     if "name" not in request.data:
    #         raise ParseError("Request to add required parameters")
    #         return Response(status=status.HTTP_400_BAD_REQUEST)

    #     filename = "temp_doc_for_conversion.doc"  # received file name
    #     file_obj = request.data["name"]
    #     with default_storage.open("jd_temp/" + filename, "wb+") as destination:
    #         for chunk in file_obj.chunks():
    #             destination.write(chunk)
    #     dirName = os.path.dirname(__file__)
    #     cwd = Path.cwd()
    #     filename = os.path.join(cwd, "media", "jd_temp/", "temp_doc_for_conversion.doc")
    #     texts = extract_text_from_doc(filename)
    #     return Response(texts, status=status.HTTP_200_OK)

    # def post(self, request, format=None):
    #     if "name" not in request.data:
    #         raise ParseError("Request to add required parameters")

    #     filename = "temp_pdf_for_conversion.pdf"  # received file name
    #     file_obj = request.data["name"]
    #     with default_storage.open("jd_temp/" + filename, "wb+") as destination:
    #         for chunk in file_obj.chunks():
    #             destination.write(chunk)
    #     dirName = os.path.dirname(__file__)
    #     cwd = Path.cwd()
    #     filename = os.path.join(cwd, "media", "jd_temp/", "temp_pdf_for_conversion.pdf")
    #     texts = extract_text_from_pdf_with_image(filename)
    #     return Response(texts, status=status.HTTP_200_OK)
