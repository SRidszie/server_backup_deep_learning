import time
from rest_framework import response, status
import re
from django.http.response import JsonResponse
from django.http import HttpResponse

# import MySQLdb
import pytesseract
from PIL import Image, ImageFilter
import requests
import docx2txt
import pdf2image
import PyPDF2
import pyttsx3
import os
import textract
import spacy
import json
import logging
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from io import StringIO

import io
import base64
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
from docx import Document

import json, requests
import pandas as pd
from pandas.io.json import json_normalize
import jsonpickle
from json import JSONEncoder


# DF conversion of output
import pandas as pd
from dframcy import DframCy


# Set URL
url = "http://127.0.0.1:8000/jd-parse/"  # for local
###########################################################


# url_for_domain = "https://devparse.mycareercube.com/jd-parse/"


nlp = spacy.load("en_jd_parser")
dframcy = DframCy(nlp)

# data_path = "C:/Users/admin/Documents/resume_parsing/DeepLearningParser/data"


# def text2doc(sw):
#     start = time.time()
#     with open(os.path.join(data_path, "test.doc"), "wb") as fp:
#         fp.write(sw)
#     text = textract.process(
#         "C:/Users/admin/Documents/resume_parsing/DeepLearningParser/data/test.doc"
#     )

#     # return text
#     ## Converting Spacy output into pandas dataframe
#     text_df = dframcy.nlp(str(text))
#     token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
#         text_df, separate_entity_dframe=True
#     )
#     ## Rearrange Column Names
#     df = entity_text_dataframe[["ent_label", "ent_text"]]

#     ## Renaming Columns as per business need:

#     df["ent_label"] = df["ent_label"].replace(
#         {
#             "TTL": "j_title",
#             "JD_TY": "jd_type",
#             "JD_PRTY": "jd_priority",
#             "JD_CAT": "jd_category",
#             "CLT_TY": "client_type",
#             "ENT_NAM": "entity_id",
#             "EMP_TY": "employment_type",
#             "Job_LOC_C": "workzone_id",
#             "DEP": "functional_area_id",
#             "MAN_SK": "key_skills_and",
#             "OPT_SK": "key_skills_or",
#             "SK_MT": "description",
#             "CRT_CV": "num_of_certify_cv",
#             "TOT_CV": "num_of_cv",
#             "VAC": "jd_openings",
#             "EXP_QUAL": "expertise_qualification",
#             "ROL_RES": "roles_responsibilities",
#             "SAL_MN": "salary_min",
#             "SAL_MX": "salary_max",
#             "CLT_NM": "client_id",
#             "EDU": "edu_qualification",
#             "CNT_PRD": "contact_period",
#             "SFT_SK": "soft_skills",
#             "GEN": "gender",
#             "CTRY": "country_id",
#             "JON_DT": "joining_date",
#             "SHFT_TM": "shift_timings",
#             "PH": "contact_no",
#         }
#     )

#     df_ = (
#         df.groupby("ent_label", sort=True)["ent_text"]
#         .apply(set)
#         .reset_index(name="ent_text")
#     )

#     ## dataframe to dict

#     _df_ = df_.set_index("ent_label").T.to_dict("list")

#     ## Adding Accuracy Matrix in API:

#     # Opening JSON file
#     f = open("metrics.json", "r")

#     # returns JSON object as
#     # a dictionary
#     df_json = json.loads(f.read())

#     # print(df_json)

#     ## Adding Status Code in the API

#     ## For time Calculation - Run this at last
#     script_time = time.time() - start
#     return (
#         _df_,
#         ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
#         ("JD_Parser Model Accuracy:", df_json),
#     )


def extract_text_from_doc(doc_path):
    start = time.time()
    temp = textract.process(str(doc_path)).decode("utf-8")
    text = [line.replace("\t", " ") for line in temp.split("\n") if line]

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(str(text))
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )
    ## Rearrange Column Names
    df = entity_text_dataframe[["ent_label", "ent_text"]]

    ## Renaming Columns as per business need:

    df["ent_label"] = df["ent_label"].replace(
        {
            "TTL": "jd_title",
            "JD_TY": "jd_type",
            "JD_PRTY": "jd_priority",
            "JD_CAT": "jd_category",
            "CLT_TY": "client_type",
            "ENT_NAM": "entity_id",
            "EMP_TY": "employment_type",
            "Job_LOC_C": "comp_city",
            "DEP": "functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "num_of_certify_cv",
            "TOT_CV": "num_of_cv",
            "VAC": "jd_openings",
            "EXP_QUAL": "expertise_qualification",
            "ROL_RES": "roles_responsibilities",
            "SAL_MN": "salary_min",
            "PH": "contact_no",
            "EXP_MIN": "exp_min",
            "EXP_MAX": "exp_max",
            "EMAIL": "email",
            "CTRY": "workzone_id",
            "date": "date",  # new label
            "money": "money",  # new label
            "CLT_NAM": "client_name",  # new label
            "AGE": "age",  # new label
            "ADD": "address",  # new label
            "WEB": "website",  # new label
            "LANG": "language",  # new lable
            "CO_NAM": "compnay_name",  # new label
            "SAL_MX": "salary_max",  # new label
            "EDU": "edu_qualification",  # new label
            "CNT_PRD": "contact_period",  # new label
            "SFT_SK": "soft_skills",  # new label
            "GEN": "gender",  # new label
            "JON_DT": "joining_date",  # new label
            "SHFT_TM": "shift_timings",  # new label
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(list)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    blank_dict = {
        "jd_title": None,
        "jd_type": None,
        "jd_priority": None,
        "jd_category": None,
        "client_type": None,
        "entity_id": None,
        "employment_type": None,
        "comp_city": None,
        "functional_area_id": None,
        "key_skills_and": None,
        "key_skills_or": None,
        "description": None,
        "num_of_certify_cv": None,
        "num_of_cv": None,
        "jd_openings": None,
        "expertise_qualification": None,
        "roles_responsibilities": None,
        "salary_min": None,
        "contact_no": None,
        "exp_min": None,
        "exp_max": None,
        "email": None,
        "workzone_id": None,
        "date": None,  # new label
        "money": None,  # new label
        "client_name": None,  # new label
        "age": None,  # new label
        "address": None,  # new label
        "website": None,  # new label
        "language": None,  # new lable
        "compnay_name": None,  # new label
        "salary_max": None,  # new label
        "edu_qualification": None,  # new label
        "contact_period": None,  # new label
        "soft_skills": None,  # new label
        "gender": None,  # new label
        "joining_date": None,  # new label
        "shift_timings": None,  # new label
    }
    # print(blank_dict)

    ############################ XXXXXXXXXXXXXXXX ##############################

    merge_dict_output = blank_dict.copy()
    for key, value in _df_.items():
        merge_dict_output[key] = value

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    ##################################################

    sampleJson = jsonpickle.encode(merge_dict_output)
    decodedSet = jsonpickle.decode(sampleJson)
    manual_text_add = {
        "status": "success",
        "status_code": status.HTTP_200_OK,
        "message": "JD file Uploaded Successfully",
        # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
        "data": decodedSet,
    }
    r = json.dumps(manual_text_add, indent=10)
    loaded_r = json.loads(r)
    return loaded_r
    ###################################### commenting as of now ###############################
    # ## Adding Accuracy Matrix in API:

    # import json

    # # Opening JSON file
    # f = open("metrics.json", "r")

    # # returns JSON object as
    # # a dictionary
    # df_json = json.loads(f.read())

    # ## Adding Status Code in the API

    # ## For time Calculation - Run this at last
    # script_time = time.time() - start
    # return (
    #     _df_,
    #     ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
    #     ("JD_Parser Model Accuracy:", df_json),
    # )
    ###########################################################################################################


def extract_text_from_pdf_with_image(pdf_file):
    start = time.time()
    final_text = []
    images = pdf2image.convert_from_path(str(pdf_file))
    for pages, img in enumerate(images):
        text = pytesseract.image_to_string(img)

        final_text.append({pages, text})

    ## Spacy - NLP

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(str(final_text))
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )

    ## Rearrange Column Names
    df = entity_text_dataframe[["ent_label", "ent_text"]]

    ## Renaming Columns as per business need:

    df["ent_label"] = df["ent_label"].replace(
        {
            "TTL": "jd_title",
            "JD_TY": "jd_type",
            "JD_PRTY": "jd_priority",
            "JD_CAT": "jd_category",
            "CLT_TY": "client_type",
            "ENT_NAM": "entity_id",
            "EMP_TY": "employment_type",
            "Job_LOC_C": "comp_city",
            "DEP": "functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "num_of_certify_cv",
            "TOT_CV": "num_of_cv",
            "VAC": "jd_openings",
            "EXP_QUAL": "expertise_qualification",
            "ROL_RES": "roles_responsibilities",
            "SAL_MN": "salary_min",
            "PH": "contact_no",
            "EXP_MIN": "exp_min",
            "EXP_MAX": "exp_max",
            "EMAIL": "email",
            "CTRY": "workzone_id",
            "date": "date",  # new label
            "money": "money",  # new label
            "CLT_NAM": "client_name",  # new label
            "AGE": "age",  # new label
            "ADD": "address",  # new label
            "WEB": "website",  # new label
            "LANG": "language",  # new lable
            "CO_NAM": "compnay_name",  # new label
            "SAL_MX": "salary_max",  # new label
            "EDU": "edu_qualification",  # new label
            "CNT_PRD": "contact_period",  # new label
            "SFT_SK": "soft_skills",  # new label
            "GEN": "gender",  # new label
            "JON_DT": "joining_date",  # new label
            "SHFT_TM": "shift_timings",  # new label
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(list)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    blank_dict = {
        "jd_title": None,
        "jd_type": None,
        "jd_priority": None,
        "jd_category": None,
        "client_type": None,
        "entity_id": None,
        "employment_type": None,
        "comp_city": None,
        "functional_area_id": None,
        "key_skills_and": None,
        "key_skills_or": None,
        "description": None,
        "num_of_certify_cv": None,
        "num_of_cv": None,
        "jd_openings": None,
        "expertise_qualification": None,
        "roles_responsibilities": None,
        "salary_min": None,
        "contact_no": None,
        "exp_min": None,
        "exp_max": None,
        "email": None,
        "workzone_id": None,
        "date": None,  # new label
        "money": None,  # new label
        "client_name": None,  # new label
        "age": None,  # new label
        "address": None,  # new label
        "website": None,  # new label
        "language": None,  # new lable
        "compnay_name": None,  # new label
        "salary_max": None,  # new label
        "edu_qualification": None,  # new label
        "contact_period": None,  # new label
        "soft_skills": None,  # new label
        "gender": None,  # new label
        "joining_date": None,  # new label
        "shift_timings": None,  # new label
    }
    # print(blank_dict)

    ############################ XXXXXXXXXXXXXXXX ##############################

    merge_dict_output = blank_dict.copy()
    for key, value in _df_.items():
        merge_dict_output[key] = value

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    ##################################################

    sampleJson = jsonpickle.encode(merge_dict_output)
    decodedSet = jsonpickle.decode(sampleJson)
    manual_text_add = {
        "status": "success",
        "status_code": status.HTTP_200_OK,
        "message": "JD file Uploaded Successfully",
        # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
        "data": decodedSet,
    }
    r = json.dumps(manual_text_add, indent=10)
    loaded_r = json.loads(r)
    return loaded_r
    ###################################### commenting as of now ###############################
    # ## Adding Accuracy Matrix in API:

    # import json

    # # Opening JSON file
    # f = open("metrics.json", "r")

    # # returns JSON object as
    # # a dictionary
    # df_json = json.loads(f.read())

    # ## Adding Status Code in the API

    # ## For time Calculation - Run this at last
    # script_time = time.time() - start
    # return (
    #     _df_,
    #     ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
    #     ("JD_Parser Model Accuracy:", df_json),
    # )
    ###########################################################################################################


# def extract_text_from_image(image_file):
#     start = time.time()
#     image = Image.open(image_file)
#     try:
#         image.filter(ImageFilter.SHARPEN)
#     except ValueError:
#         print("Got an image that failed to sharpen", image_file)
#         pass
#     text = pytesseract.image_to_string(image)
#     text = text.replace("\n", " ")


def extract_text_from_image(image_file):
    start = time.time()
    image_file = base64.b64decode(image_file.encode("UTF-8"))
    buf = io.BytesIO(image_file)
    image = Image.open(buf)
    image.filter(ImageFilter.SHARPEN)
    text = pytesseract.image_to_string(image)
    text = text.replace("\n", " ")
    # print(text)
    ## Spacy - NLP

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(text)
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )

    ## Rearrange Column Names
    df = entity_text_dataframe[["ent_label", "ent_text"]]

    # custom_df = {"TTL": df.TTL, }
    ## Renaming Columns as per business need:

    df["ent_label"] = df["ent_label"].replace(
        {
            "TTL": "jd_title",
            "JD_TY": "jd_type",
            "JD_PRTY": "jd_priority",
            "JD_CAT": "jd_category",
            "CLT_TY": "client_type",
            "ENT_NAM": "entity_id",
            "EMP_TY": "employment_type",
            "Job_LOC_C": "comp_city",
            "DEP": "functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "num_of_certify_cv",
            "TOT_CV": "num_of_cv",
            "VAC": "jd_openings",
            "EXP_QUAL": "expertise_qualification",
            "ROL_RES": "roles_responsibilities",
            "SAL_MN": "salary_min",
            "PH": "contact_no",
            "EXP_MIN": "exp_min",
            "EXP_MAX": "exp_max",
            "EMAIL": "email",
            "CTRY": "workzone_id",
            "date": "date",  # new label
            "money": "money",  # new label
            "CLT_NAM": "client_name",  # new label
            "AGE": "age",  # new label
            "ADD": "address",  # new label
            "WEB": "website",  # new label
            "LANG": "language",  # new lable
            "CO_NAM": "compnay_name",  # new label
            "SAL_MX": "salary_max",  # new label
            "EDU": "edu_qualification",  # new label
            "CNT_PRD": "contact_period",  # new label
            "SFT_SK": "soft_skills",  # new label
            "GEN": "gender",  # new label
            "JON_DT": "joining_date",  # new label
            "SHFT_TM": "shift_timings",  # new label
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(list)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    blank_dict = {
        "jd_title": None,
        "jd_type": None,
        "jd_priority": None,
        "jd_category": None,
        "client_type": None,
        "entity_id": None,
        "employment_type": None,
        "comp_city": None,
        "functional_area_id": None,
        "key_skills_and": None,
        "key_skills_or": None,
        "description": None,
        "num_of_certify_cv": None,
        "num_of_cv": None,
        "jd_openings": None,
        "expertise_qualification": None,
        "roles_responsibilities": None,
        "salary_min": None,
        "contact_no": None,
        "exp_min": None,
        "exp_max": None,
        "email": None,
        "workzone_id": None,
        "date": None,  # new label
        "money": None,  # new label
        "client_name": None,  # new label
        "age": None,  # new label
        "address": None,  # new label
        "website": None,  # new label
        "language": None,  # new lable
        "compnay_name": None,  # new label
        "salary_max": None,  # new label
        "edu_qualification": None,  # new label
        "contact_period": None,  # new label
        "soft_skills": None,  # new label
        "gender": None,  # new label
        "joining_date": None,  # new label
        "shift_timings": None,  # new label
    }
    # print(blank_dict)

    ############################ XXXXXXXXXXXXXXXX ##############################

    merge_dict_output = blank_dict.copy()
    for key, value in _df_.items():
        merge_dict_output[key] = value

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    ##################################################

    ##################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################
    # Extracting only Numeric values from the dict(values) e.g
    # 1) exp_min, 2) exp_max, 3) salary_min, 4) salary_max, 5) num_of_certify_cv, 6) num_of_cv, 7) jd_openings, 8) age, 9) money, 10) contact_no

    # 1) for "exp_min"  -->
    if merge_dict_output["exp_min"] is None:
        pass
    else:
        merge_dict_output["exp_min"] = "".join(
            str(e) for e in merge_dict_output["exp_min"]
        )
        exp_min = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["exp_min"]))):
            exp_min += i

        merge_dict_output["exp_min"] = exp_min

    # # 2) for "exp_max"  -->
    if merge_dict_output["exp_max"] is None:
        pass
    else:
        merge_dict_output["exp_max"] = "".join(
            str(e) for e in merge_dict_output["exp_max"]
        )
        exp_max = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["exp_max"]))):
            exp_max += i

        merge_dict_output["exp_max"] = exp_max

    # 3) for "salary_min"  -->
    if merge_dict_output["salary_min"] is None:
        pass
    else:
        merge_dict_output["salary_min"] = "".join(
            str(e) for e in merge_dict_output["salary_min"]
        )
        salary_min = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary_min"]))):
            salary_min += i

        merge_dict_output["salary_min"] = salary_min

    # 4) for "salary_max"  -->
    if merge_dict_output["salary_max"] is None:
        pass
    else:
        merge_dict_output["salary_max"] = "".join(
            str(e) for e in merge_dict_output["salary_max"]
        )
        salary_max = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary_max"]))):
            salary_max += i

        merge_dict_output["salary_max"] = salary_max

    # 5) for "num_of_certify_cv"  -->
    if merge_dict_output["num_of_certify_cv"] is None:
        pass
    else:
        merge_dict_output["num_of_certify_cv"] = "".join(
            str(e) for e in merge_dict_output["num_of_certify_cv"]
        )
        num_of_certify_cv = int()
        for i in list(
            map(int, re.findall(r"\d+", merge_dict_output["num_of_certify_cv"]))
        ):
            num_of_certify_cv += i

        merge_dict_output["num_of_certify_cv"] = num_of_certify_cv

    # 6) for "num_of_cv"  -->
    if merge_dict_output["num_of_cv"] is None:
        pass
    else:
        merge_dict_output["num_of_cv"] = "".join(
            str(e) for e in merge_dict_output["num_of_cv"]
        )
        num_of_cv = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["num_of_cv"]))):
            num_of_cv += i

        merge_dict_output["num_of_cv"] = num_of_cv

    # 7) for "jd_openings"  -->
    if merge_dict_output["jd_openings"] is None:
        pass
    else:
        merge_dict_output["jd_openings"] = "".join(
            str(e) for e in merge_dict_output["jd_openings"]
        )
        jd_openings = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["jd_openings"]))):
            jd_openings += i

        merge_dict_output["jd_openings"] = jd_openings

    # 8) for "age"  -->
    if merge_dict_output["age"] is None:
        pass
    else:
        merge_dict_output["age"] = "".join(str(e) for e in merge_dict_output["age"])
        age = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["age"]))):
            age += i

        merge_dict_output["age"] = age

    # 9) for "money"  -->
    if merge_dict_output["money"] is None:
        pass
    else:
        merge_dict_output["money"] = "".join(str(e) for e in merge_dict_output["money"])
        money = int()
        for i in list(map(int, re.findall(r"\d+", merge_dict_output["money"]))):
            money += i

        merge_dict_output["money"] = money

    # 10) for "contact_no"  -->
    # if merge_dict_output["contact_no"] is None:
    #     pass
    # else:
    #     merge_dict_output["contact_no"] = "".join(
    #         str(e) for e in merge_dict_output["contact_no"]
    #     )
    #     contact_no = int()
    #     for i in list(map(int, re.findall(r"\d+", merge_dict_output["contact_no"]))):
    #         contact_no += i

    #     merge_dict_output["contact_no"] = contact_no

    ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

    # Converting dict to json by serialization -->

    sampleJson = jsonpickle.encode(merge_dict_output)
    decodedSet = jsonpickle.decode(sampleJson)
    manual_text_add = {
        "status": "success",
        "status_code": status.HTTP_200_OK,
        "message": "JD file Uploaded Successfully",
        # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
        "data": decodedSet,
    }
    r = json.dumps(manual_text_add, indent=10)
    loaded_r = json.loads(r)
    return loaded_r
    ###################################### commenting as of now ###############################
    # ## Adding Accuracy Matrix in API:

    # import json

    # # Opening JSON file
    # f = open("metrics.json", "r")

    # # returns JSON object as
    # # a dictionary
    # df_json = json.loads(f.read())

    # ## Adding Status Code in the API

    # ## For time Calculation - Run this at last
    # script_time = time.time() - start
    # return (
    #     _df_,
    #     ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
    #     ("JD_Parser Model Accuracy:", df_json),
    # )
    ###########################################################################################################


def extract_text_from_image_path(image_file_path):
    with open(image_file_path, "rb") as image_file:
        return extract_text_from_image(image_file)


# def extract_text_from_docx(doc_path):
#     start = time.time()
#     """
#     Helper function to extract plain text from .docx files

#     :param doc_path: path to .docx file to be extracted
#     :return: string of extracted text

#     """

#     try:
#         doc_path = base64.b64decode(doc_path.encode("UTF-8"))
#         temp = docx2txt.process(doc_path)
#         # fn - imput - base64 input str  / not
#         text = [line.replace("\t", " ") for line in temp.split("\n") if line]


def extract_text_from_docx(doc_path):
    start = time.time()
    """
    Helper function to extract plain text from .docx files

    :param doc_path: path to .docx file to be extracted
    :return: string of extracted text

    """
    try:
        doc_path = base64.b64decode(doc_path.encode("UTF-8"))
        bio_ = io.BytesIO(doc_path)
        temp = docx2txt.process(bio_)
        # fn - imput - base64 input str  / not
        text = [line.replace("\t", " ") for line in temp.split("\n") if line]
        # print(text)

        ## Spacy - NLP
        ## Converting Spacy output into pandas dataframe
        text_df = dframcy.nlp(str(text))
        token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
            text_df, separate_entity_dframe=True
        )

        ## Rearrange Column Names
        df = entity_text_dataframe[["ent_label", "ent_text"]]

        ## Renaming Columns as per business need:

        df["ent_label"] = df["ent_label"].replace(
            {
                "TTL": "jd_title",
                "JD_TY": "jd_type",
                "JD_PRTY": "jd_priority",
                "JD_CAT": "jd_category",
                "CLT_TY": "client_type",
                "ENT_NAM": "entity_id",
                "EMP_TY": "employment_type",
                "Job_LOC_C": "comp_city",
                "DEP": "functional_area_id",
                "MAN_SK": "key_skills_and",
                "OPT_SK": "key_skills_or",
                "SK_MT": "description",
                "CRT_CV": "num_of_certify_cv",
                "TOT_CV": "num_of_cv",
                "VAC": "jd_openings",
                "EXP_QUAL": "expertise_qualification",
                "ROL_RES": "roles_responsibilities",
                "SAL_MN": "salary_min",
                "PH": "contact_no",
                "EXP_MIN": "exp_min",
                "EXP_MAX": "exp_max",
                "EMAIL": "email",
                "CTRY": "workzone_id",
                "date": "date",  # new label
                "money": "money",  # new label
                "CLT_NAM": "client_name",  # new label
                "AGE": "age",  # new label
                "ADD": "address",  # new label
                "WEB": "website",  # new label
                "LANG": "language",  # new lable
                "CO_NAM": "compnay_name",  # new label
                "SAL_MX": "salary_max",  # new label
                "EDU": "edu_qualification",  # new label
                "CNT_PRD": "contact_period",  # new label
                "SFT_SK": "soft_skills",  # new label
                "GEN": "gender",  # new label
                "JON_DT": "joining_date",  # new label
                "SHFT_TM": "shift_timings",  # new label
            }
        )

        df_ = (
            df.groupby("ent_label", sort=True)["ent_text"]
            .apply(list)
            .reset_index(name="ent_text")
        )

        ## dataframe to dict

        _df_ = df_.set_index("ent_label").T.to_dict("list")

        blank_dict = {
            "jd_title": None,
            "jd_type": None,
            "jd_priority": None,
            "jd_category": None,
            "client_type": None,
            "entity_id": None,
            "employment_type": None,
            "comp_city": None,
            "functional_area_id": None,
            "key_skills_and": None,
            "key_skills_or": None,
            "description": None,
            "num_of_certify_cv": None,
            "num_of_cv": None,
            "jd_openings": None,
            "expertise_qualification": None,
            "roles_responsibilities": None,
            "salary_min": None,
            "contact_no": None,
            "exp_min": None,
            "exp_max": None,
            "email": None,
            "workzone_id": None,
            "date": None,  # new label
            "money": None,  # new label
            "client_name": None,  # new label
            "age": None,  # new label
            "address": None,  # new label
            "website": None,  # new label
            "language": None,  # new lable
            "compnay_name": None,  # new label
            "salary_max": None,  # new label
            "edu_qualification": None,  # new label
            "contact_period": None,  # new label
            "soft_skills": None,  # new label
            "gender": None,  # new label
            "joining_date": None,  # new label
            "shift_timings": None,  # new label
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################

        ##################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) exp_min, 2) exp_max, 3) salary_min, 4) salary_max, 5) num_of_certify_cv, 6) num_of_cv, 7) jd_openings, 8) age, 9) money, 10) contact_no

        # 1) for "exp_min"  -->
        if merge_dict_output["exp_min"] is None:
            pass
        else:
            merge_dict_output["exp_min"] = "".join(
                str(e) for e in merge_dict_output["exp_min"]
            )
            exp_min = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["exp_min"]))):
                exp_min += i

            merge_dict_output["exp_min"] = exp_min

        # # 2) for "exp_max"  -->
        if merge_dict_output["exp_max"] is None:
            pass
        else:
            merge_dict_output["exp_max"] = "".join(
                str(e) for e in merge_dict_output["exp_max"]
            )
            exp_max = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["exp_max"]))):
                exp_max += i

            merge_dict_output["exp_max"] = exp_max

        # 3) for "salary_min"  -->
        if merge_dict_output["salary_min"] is None:
            pass
        else:
            merge_dict_output["salary_min"] = "".join(
                str(e) for e in merge_dict_output["salary_min"]
            )
            salary_min = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["salary_min"]))
            ):
                salary_min += i

            merge_dict_output["salary_min"] = salary_min

        # 4) for "salary_max"  -->
        if merge_dict_output["salary_max"] is None:
            pass
        else:
            merge_dict_output["salary_max"] = "".join(
                str(e) for e in merge_dict_output["salary_max"]
            )
            salary_max = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["salary_max"]))
            ):
                salary_max += i

            merge_dict_output["salary_max"] = salary_max

        # 5) for "num_of_certify_cv"  -->
        if merge_dict_output["num_of_certify_cv"] is None:
            pass
        else:
            merge_dict_output["num_of_certify_cv"] = "".join(
                str(e) for e in merge_dict_output["num_of_certify_cv"]
            )
            num_of_certify_cv = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["num_of_certify_cv"]))
            ):
                num_of_certify_cv += i

            merge_dict_output["num_of_certify_cv"] = num_of_certify_cv

        # 6) for "num_of_cv"  -->
        if merge_dict_output["num_of_cv"] is None:
            pass
        else:
            merge_dict_output["num_of_cv"] = "".join(
                str(e) for e in merge_dict_output["num_of_cv"]
            )
            num_of_cv = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["num_of_cv"]))):
                num_of_cv += i

            merge_dict_output["num_of_cv"] = num_of_cv

        # 7) for "jd_openings"  -->
        if merge_dict_output["jd_openings"] is None:
            pass
        else:
            merge_dict_output["jd_openings"] = "".join(
                str(e) for e in merge_dict_output["jd_openings"]
            )
            jd_openings = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["jd_openings"]))
            ):
                jd_openings += i

            merge_dict_output["jd_openings"] = jd_openings

        # 8) for "age"  -->
        if merge_dict_output["age"] is None:
            pass
        else:
            merge_dict_output["age"] = "".join(str(e) for e in merge_dict_output["age"])
            age = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["age"]))):
                age += i

            merge_dict_output["age"] = age

        # 9) for "money"  -->
        if merge_dict_output["money"] is None:
            pass
        else:
            merge_dict_output["money"] = "".join(
                str(e) for e in merge_dict_output["money"]
            )
            money = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["money"]))):
                money += i

            merge_dict_output["money"] = money

        # 10) for "contact_no"  -->
        # if merge_dict_output["contact_no"] is None:
        #     pass
        # else:
        #     merge_dict_output["contact_no"] = "".join(
        #         str(e) for e in merge_dict_output["contact_no"]
        #     )
        #     contact_no = int()
        #     for i in list(map(int, re.findall(r"\d+", merge_dict_output["contact_no"]))):
        #         contact_no += i

        #     merge_dict_output["contact_no"] = contact_no

        ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "JD file Uploaded Successfully",
            # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
        ###################################### commenting as of now ###############################
        # ## Adding Accuracy Matrix in API:

        # import json

        # # Opening JSON file
        # f = open("metrics.json", "r")

        # # returns JSON object as
        # # a dictionary
        # df_json = json.loads(f.read())

        # ## Adding Status Code in the API

        # ## For time Calculation - Run this at last
        # script_time = time.time() - start
        # return (
        #     _df_,
        #     ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
        #     ("JD_Parser Model Accuracy:", df_json),
        # )
        ###########################################################################################################

    except KeyError:
        error_status_code_docx = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_docx


def extract_text_from_pdf(pdf_file):
    start = time.time()
    """
    A utility function to convert a machine-readable PDF to raw text.
    This code is largely borrowed from existing solutions, and does not match the style of the rest of this repo.
    :param input_pdf_path: Path to the .pdf file which should be converted
    :type input_pdf_path: str
    :return: The text contents of the pdf
    :rtype: str
    """
    try:

        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        codec = "utf-8"
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        # page_index=3

        # Iterate through pages
        for page in PDFPage.get_pages(
            pdf_file,
            set(),
            maxpages=0,
            password="",
            caching=True,
            check_extractable=True,
        ):
            interpreter.process_page(page)

        device.close()

        # Get full string from PDF
        text = retstr.getvalue()
        retstr.close()
        doc = nlp(text)
        ## NLP Output in DataFrame format - directly using dframcy (spacy add on)
        # for separate entity_text dataframe
        text_df = dframcy.nlp(text)
        token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
            text_df, separate_entity_dframe=True
        )

        ## Rearrange Column Names
        df = entity_text_dataframe[["ent_label", "ent_text"]]

        ## Renaming Columns as per business need:

        df["ent_label"] = df["ent_label"].replace(
            {
                "TTL": "jd_title",
                "JD_TY": "jd_type",
                "JD_PRTY": "jd_priority",
                "JD_CAT": "jd_category",
                "CLT_TY": "client_type",
                "ENT_NAM": "entity_id",
                "EMP_TY": "employment_type",
                "Job_LOC_C": "comp_city",
                "DEP": "functional_area_id",
                "MAN_SK": "key_skills_and",
                "OPT_SK": "key_skills_or",
                "SK_MT": "description",
                "CRT_CV": "num_of_certify_cv",
                "TOT_CV": "num_of_cv",
                "VAC": "jd_openings",
                "EXP_QUAL": "expertise_qualification",
                "ROL_RES": "roles_responsibilities",
                "SAL_MN": "salary_min",
                "PH": "contact_no",
                "EXP_MIN": "exp_min",
                "EXP_MAX": "exp_max",
                "EMAIL": "email",
                "CTRY": "workzone_id",
                "date": "date",  # new label
                "money": "money",  # new label
                "CLT_NAM": "client_name",  # new label
                "AGE": "age",  # new label
                "ADD": "address",  # new label
                "WEB": "website",  # new label
                "LANG": "language",  # new lable
                "CO_NAM": "compnay_name",  # new label
                "SAL_MX": "salary_max",  # new label
                "EDU": "edu_qualification",  # new label
                "CNT_PRD": "contact_period",  # new label
                "SFT_SK": "soft_skills",  # new label
                "GEN": "gender",  # new label
                "JON_DT": "joining_date",  # new label
                "SHFT_TM": "shift_timings",  # new label
            }
        )

        df_ = (
            df.groupby("ent_label", sort=True)["ent_text"]
            .apply(list)
            .reset_index(name="ent_text")
        )

        ## dataframe to dict

        _df_ = df_.set_index("ent_label").T.to_dict("list")

        blank_dict = {
            "jd_title": None,
            "jd_type": None,
            "jd_priority": None,
            "jd_category": None,
            "client_type": None,
            "entity_id": None,
            "employment_type": None,
            "comp_city": None,
            "functional_area_id": None,
            "key_skills_and": None,
            "key_skills_or": None,
            "description": None,
            "num_of_certify_cv": None,
            "num_of_cv": None,
            "jd_openings": None,
            "expertise_qualification": None,
            "roles_responsibilities": None,
            "salary_min": None,
            "contact_no": None,
            "exp_min": None,
            "exp_max": None,
            "email": None,
            "workzone_id": None,
            "date": None,  # new label
            "money": None,  # new label
            "client_name": None,  # new label
            "age": None,  # new label
            "address": None,  # new label
            "website": None,  # new label
            "language": None,  # new lable
            "compnay_name": None,  # new label
            "salary_max": None,  # new label
            "edu_qualification": None,  # new label
            "contact_period": None,  # new label
            "soft_skills": None,  # new label
            "gender": None,  # new label
            "joining_date": None,  # new label
            "shift_timings": None,  # new label
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "JD file Uploaded Successfully",
            # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
        ###################################### commenting as of now ###############################
        # ## Adding Accuracy Matrix in API:

        # import json

        # # Opening JSON file
        # f = open("metrics.json", "r")

        # # returns JSON object as
        # # a dictionary
        # df_json = json.loads(f.read())

        # ## Adding Status Code in the API

        # ## For time Calculation - Run this at last
        # script_time = time.time() - start
        # return (
        #     _df_,
        #     ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
        #     ("JD_Parser Model Accuracy:", df_json),
        # )
        ###########################################################################################################

    except Exception as e:
        logging.error("Error in pdf file:: " + str(e))
        return []


def extract_content_filetype(file_content, file_type):
    """
    Wrapper function to detect the file extension and call text
    extraction function accordingly

    :param file_path: path of file of which text is to be extracted
    :param extension: extension of file `file_name`
    """
    jd_lines = ""
    if file_type == ".docx":
        jd_lines = extract_text_from_docx(file_content)
        return jd_lines
    elif file_type == ".doc":
        jd_lines = extract_text_from_doc(file_content)
        return jd_lines
    elif file_type == ".pdf":
        jd_lines = extract_text_from_pdf(file_content)
        return jd_lines
    elif file_type == ".png":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_type == ".PNG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_type == ".jpg":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_type == ".JPG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_type == ".jpeg":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_type == ".JPEG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    else:
        error_status_code = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code


# def process(file):
#     # import pdb
#     # pdb.set_trace()
#     """
#     Main function to process resume file to json.
#     :param file: Resume file
#     :return: resume_data: Parsed resume dictionary
#     """
#     if file.name.endswith("docx"):
#         jd_lines = extract_text_from_docx(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("doc"):
#         jd_lines = extract_text_from_doc(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("PNG"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("png"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("JPG"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("jpg"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("jpeg"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("JPEG"):
#         jd_lines = extract_text_from_image(file)
#         # print(jd_lines)
#         return jd_lines
#     elif file.name.endswith("pdf"):
#         jd_lines = extract_text_from_pdf(file)
#         # print(jd_lines)
#         return jd_lines
#     else:
#         return None
