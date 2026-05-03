from django.shortcuts import render

# Create your views here.
from django.shortcuts import render, redirect
from jd_parser.forms import ImageFileForm
from jd_parser.models import ImageFile
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
# from rest_framework_simplejwt.authentication import JWTAuthentication
from rest_framework.authentication import TokenAuthentication
import base64
from django.http.response import JsonResponse
from rest_framework.permissions import IsAuthenticated
from django.http import HttpResponse
from rest_framework.parsers import JSONParser

# end

pytesseract.pytesseract.tesseract_cmd = (
    # "C:/Program Files/Tesseract-OCR/tesseract.exe"  # your path may be different
    "/usr/bin/tesseract"
)


from django.conf import settings
from rest_framework.exceptions import ParseError
from rest_framework.parsers import BaseParser

class PlainTextParser(BaseParser):
    media_type = "text/plain"

    def parse(self, stream, media_type=None, parser_context=None):
        """
        Parses the incoming bytestream as Plain Text and returns the resulting data.
        """
        parser_context = parser_context or {}
        encoding = parser_context.get('encoding', settings.DEFAULT_CHARSET)

        try:
            codecs=" "
            decoded_stream = codecs.getreader(encoding)(stream)
            text_content = decoded_stream.read()
            return text_content
        except ValueError as exc:
            raise ParseError('Plain text parse error - %s' % str(exc)) 
        

    

class cv_search_api(APIView):
    permission_classes = (IsAuthenticated,)
    parser_class = (PlainTextParser,JSONParser)

    def post(self, request):
        if "keyword_data"  not in request.data:
        # if "keyword_data" and "cv_data" not in request.data:
        
            response = JsonResponse(
                {
                    "status": "failure",
                    "status-code": status.HTTP_400_BAD_REQUEST,
                    "message": "Request to add required parameters",
                }
            )
            return response
        else:
            response = JsonResponse(
                {
                    "status": "Check Parameter Key",
                    "status-code": status.HTTP_400_BAD_REQUEST,
                    "message": "Request to add required keword parameters Received",
                }
            )   

        keyword_data = request.data["keyword_data"]
        print(keyword_data)
 
        # cv_data = request.data["cv_data"]
        # print(cv_data)
        # return HttpResponse(keyword_data,cv_data)

        # text = cv_search_from_text(keyword_data, cv_data)
        text = cv_search(keyword_data)
        return Response(text)





