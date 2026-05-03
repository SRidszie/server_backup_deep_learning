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

from django.views.generic import RedirectView
from django.views.decorators.csrf import csrf_exempt
from .views import *

urlpatterns = [
 
    path("pdf/", pdf_with_image_to_view.as_view(), name="pdf"),
    path("api/", img_pdf_docx_to_json_view.as_view(), name="api"),
    path("doc/", doc_without_image_to_view.as_view(), name="doc"),
    
]
