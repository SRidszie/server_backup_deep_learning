from django.contrib import admin
from jd_parser.models import *


class JDAdmin(admin.ModelAdmin): 
    list_display = ('jd_title', 'jd_type', 'jd_priority','jd_category','client_type','entity','No_Of_Openings','Department','Salary') 



admin.site.register(ImageFile)

