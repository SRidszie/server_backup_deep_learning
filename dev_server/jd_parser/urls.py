"""server URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.urls import path
from jd_parser.views import *

# from jd_parser.views import PostViews
from django.views.generic import RedirectView
from django.views.decorators.csrf import csrf_exempt
from .views import *

urlpatterns = [
    # path("upload", home, name="home"),
    # path('doc/',upload_doc,name='upload_doc'),
    # path("jd-parse/pdf/", pdf_with_image_to_view.as_view(), name="pdf"),
    # path("jd-parse/", img_pdf_docx_to_json_view.as_view(), name="api"),
    path("jd-parse/", image_doc_docx_pdf_txt_rtf_to_json_view.as_view(), name="api"),
    # path("jd-parse/doc/", doc_without_image_to_view.as_view(), name="doc"),
    # path("jd/", JD_objects.as_view()),
    # path("jd/<int:pk>/", JD_objects_details.as_view()),
    # path('', RedirectView.as_view(url='/ocr/url/', permanent=True)),
    # path("jd-parse/image/", OcrDataApiView.as_view(), name="ocr_image"),
    # path('api/v1/ocr/doc/', OcrDOCSUrlApiView.as_view(), name='ocr_doc'),
    # path("json/", OcrUrlApiView.as_view(), name="json"),
    # path('ocr/url/', OcrUrlView.as_view(), name='ocr_url'),
    # path('', PostViews.as_view(), name='file-upload'),
]
