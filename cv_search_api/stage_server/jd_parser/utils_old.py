import time
from rest_framework import response, status
import re
from django.http.response import JsonResponse
from django.http import HttpResponse
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

from striprtf.striprtf import rtf_to_text

# Set URL
url = "http://127.0.0.1:8000/jd-parse/"  # for local
###########################################################

nlp = spacy.load("en_jd_parser")
dframcy = DframCy(nlp)


############################################ Word Document Fn ################################################################

################################################ .doc ########################################################################
def extract_text_from_doc(doc_path):
    start = time.time()
    try:
        # doc_path = base64.b64decode(doc_path.encode("UTF-8"))
        # bio_ = io.BytesIO(doc_path)
        # temp = textract.process(str(bio_))
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
        error_status_code_doc = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_doc


################################################ .docx ################################################################################


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
        # temp = docx2txt.process(doc_path)
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


######################################### .doc with image #####################################################################


def extract_text_from_doc_with_image(doc_file):
    start = time.time()
    try:
        output_file = (
            os.path.splitext(str(doc_file))[0] + ".pdf"
        )  # adding pdf extension to doc file
        # output_file = base64.b64decode(output_file.encode("UTF-8"))
        # bio_ = io.BytesIO(output_file)

        final_text = []
        # str(doc_path)).decode("utf-8")
        images = pdf2image.convert_from_path(output_file)
        for pages, img in enumerate(images):
            text = pytesseract.image_to_string(img)

            final_text.append({pages, text})
        final_text = [
            line.replace("\t", " ") for line in str(final_text).split("\n") if line
        ]

        # return final_text
        ## Converting Spacy output into pandas dataframe
        text_df = dframcy.nlp(str(final_text))
        # return text_df
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
        error_status_code_doc_with_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_doc_with_image


###################################################### .pdf ##############################################################################


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
        pdf_file = base64.b64decode(pdf_file.encode("UTF-8"))
        bio_ = io.BytesIO(pdf_file)
        for page in PDFPage.get_pages(
            bio_,
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
        # doc = nlp(text)
        # pdf_path = base64.b64decode(pdf_file.encode("UTF-8"))
        # bio_ = io.BytesIO(pdf_path)
        # text = pdf_to_text(bio_)
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
        error_status_code_pdf = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_pdf


##################################################### .pdf with image ####################################################################


def extract_text_from_pdf_with_image(pdf_file):
    start = time.time()
    try:
        pdf_file = base64.b64decode(pdf_file.encode("UTF-8"))
        bio_ = io.BytesIO(pdf_file)

        final_text = []
        images = pdf2image.convert_from_path(str(bio_))
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
        error_status_code_pdf_with_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_pdf_with_image


############################################### Image Fn (.jpg, .png, .jpeg) ###################################################################


def extract_text_from_image(image_file):
    start = time.time()
    try:
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
        error_status_code_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_image


######################################################## Excel and CSV Fn ############################################################

######################################################## .csv ########################################################################


def extract_text_from_csv(csv_path):
    start = time.time()
    try:
        csv_path = base64.b64decode(csv_path.encode("UTF-8"))
        bio_ = io.BytesIO(csv_path)
        # text_x = pd.read_csv(bio_.decode("UTF-8"))
        text_x = pd.read_csv(bio_, delimiter="utf-8")
        # print(text_x)
        # return text
        # text = pd.DataFrame.to_string(text_x)
        # print(text)
        # print(type(b))
        # return b
        # Converting Spacy output into pandas dataframe
        text_df = dframcy.nlp(str(text_x))
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
        error_status_code_csv = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_csv


#################################################### .xlsx #######################################################################

# Excel to text -->


def extract_text_from_excel(excel_path):
    start = time.time()
    try:
        excel_path = base64.b64decode(excel_path.encode("UTF-8"))
        bio_ = io.BytesIO(excel_path)
        text_x = pd.read_excel(bio_, index_col=0)
        # text_x = pd.read_csv(excel_path, delimiter="utf-8")
        # print(text_x)
        # return text
        text = pd.DataFrame.to_string(text_x)
        # print(text)
        # print(type(b))
        # return b
        # Converting Spacy output into pandas dataframe
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
        error_status_code_excel = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_excel


######################################################## Text Fn #######################################################################

######################################################## .txt ##########################################################################
def extract_text_from_text(text_path):
    start = time.time()
    try:
        text_path = base64.b64decode(text_path.encode("UTF-8"))
        bio_ = io.BytesIO(text_path)
        # return bio_
        text_x = pd.read_csv(bio_, error_bad_lines=False)

        # print(text_x)
        # text_x = pd.read_csv(text_path, delimiter="utf-8")
        # print(text_x)
        # return text
        # text = pd.DataFrame.to_string(text_x)
        # print(text)
        # print(type(b))
        # return b
        # Converting Spacy output into pandas dataframe
        text_df = dframcy.nlp(str(text_x))
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
    except:
        error_status_code_text = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_text


############################################################### .rtf #####################################################################


def extract_text_from_rtf(rtf_path):
    start = time.time()
    try:
        rtf_path = base64.b64decode(rtf_path.encode("UTF-8"))
        # bio_ = io.BytesIO(rtf_path)
        # return bio_
        text = rtf_to_text(str(rtf_path))
        # return text

        # print(text)
        # text_x = pd.read_csv(text_path, delimiter="utf-8")
        # print(text_x)
        # return text
        # text = pd.DataFrame.to_string(text_x)
        # print(text)
        # print(type(text))
        # return b
        # Converting Spacy output into pandas dataframe
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
    except:
        error_status_code_rtf = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_rtf


################################################# XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ##################################################
def process(file_content, file_extension):
    """
    Wrapper function to detect the file extension and call text
    extraction function accordingly

    :param file_path: path of file of which text is to be extracted
    :param extension: extension of file `file_name`
    """
    jd_lines = ""
    if file_extension == ".docx":
        jd_lines = extract_text_from_docx(file_content)
        return jd_lines
    elif file_extension == ".DOCX":
        jd_lines = extract_text_from_docx(file_content)
        return jd_lines
    # elif file_extension == ".doc":
    #     jd_lines = extract_text_from_doc(file_content)
    #     return jd_lines
    # elif file_extension == ".DOC":
    #     jd_lines = extract_text_from_doc(file_content)
    #     return jd_lines
    elif file_extension == ".doc":
        jd_lines = extract_text_from_doc_with_image(file_content)
        return jd_lines
    elif file_extension == ".DOC":
        jd_lines = extract_text_from_doc_with_image(file_content)
        return jd_lines
    elif file_extension == ".pdf":
        jd_lines = extract_text_from_pdf(file_content)
        return jd_lines
    elif file_extension == ".PDF":
        jd_lines = extract_text_from_pdf(file_content)
        return jd_lines
    elif file_extension == ".pdf":
        jd_lines = extract_text_from_pdf_with_image(file_content)
        return jd_lines
    elif file_extension == ".PDF":
        jd_lines = extract_text_from_pdf_with_image(file_content)
        return jd_lines
    elif file_extension == ".png":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".PNG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".jpg":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".JPG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".jpeg":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".JPEG":
        jd_lines = extract_text_from_image(file_content)
        return jd_lines
    elif file_extension == ".csv":
        jd_lines = extract_text_from_csv(file_content)
        return jd_lines
    elif file_extension == ".CSV":
        jd_lines = extract_text_from_csv(file_content)
        return jd_lines
    elif file_extension == ".xlxs":
        jd_lines = extract_text_from_excel(file_content)
        return jd_lines
    elif file_extension == ".XLXS":
        jd_lines = extract_text_from_excel(file_content)
        return jd_lines
    elif file_extension == ".txt":
        jd_lines = extract_text_from_text(file_content)
        return jd_lines
    elif file_extension == ".TXT":
        jd_lines = extract_text_from_text(file_content)
        return jd_lines
    elif file_extension == ".rtf":
        jd_lines = extract_text_from_rtf(file_content)
        return jd_lines
    elif file_extension == ".RTF":
        jd_lines = extract_text_from_rtf(file_content)
        return jd_lines
    else:
        error_status_code = {
            "status": "failure",
            "status_code": status.HTTP_422_UNPROCESSABLE_ENTITY,
            "message": "file_content and file_extension Mismatch/Missing",
        }
        return error_status_code
