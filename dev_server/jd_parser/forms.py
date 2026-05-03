from django import forms
from jd_parser.models import *


class ImageFileForm(forms.ModelForm):
    class Meta:
        model = ImageFile
        fields = ('image', )









