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
url = "http://127.0.0.1:8000/Resume-parse/"  # for local
###########################################################

nlp = spacy.load("en_resume_parser")
dframcy = DframCy(nlp)

############################################ Word Document Fn ################################################################

################################################ .doc ########################################################################


def extract_text_from_doc(doc_path):
    start = time.time()
    try:
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except:
        error_status_code_doc = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_doc
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
    #     ("Resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
    #     ("Resume_Parser Model Accuracy:", df_json),
    # )
    ###########################################################################################################


################################################ .doc with image ##########################################################################
def extract_text_from_doc_with_image(doc_file):
    start = time.time()
    try:
        output_file = os.path.splitext(str(doc_file))[0] + ".pdf"
        # wdFormatPDF = 17
        # input_file = os.path.abspath(file1)
        # output_file = os.path.splitext(doc_file)[0] + ".pdf"
        # output_file = os.path.abspath(output_file)
        # # create COM object
        # word = comtypes.client.CreateObject("Word.Application")
        # # key point 1: make word visible before open a new document
        # word.Visible = True
        # key point 2: wait for the COM Server to prepare well.
        # time.sleep(3)

        # convert docx file 1 to pdf file 1
        # doc = word.Documents.Open(doc_file)
        # doc = doc_file
        # doc.SaveAs(output_file, FileFormat=wdFormatPDF)
        # doc.Close()
        # word.Quit()

        # return output_file
        final_text = []
        images = pdf2image.convert_from_path(str(output_file))
        for pages, img in enumerate(images):
            text = pytesseract.image_to_string(img)

            final_text.append({pages, text})
        # return final_text
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start

        #################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        # 12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version
        #
        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

            # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = salary

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except KeyError:
        error_status_code_doc_with_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_doc_with_image

    ################################################# .docx ##################################################################


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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
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
        #     ("Resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
        #     ("Resume_Parser Model Accuracy:", df_json),
        # )
        ###########################################################################################################

    except KeyError:
        error_status_code_docx = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
        return error_status_code_docx


######################################################### Image Fn (.jpg, .png, .jpeg) #####################################################


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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": [["Abhishek"]],
            "last_name": [["Sharma"]],
            "Email": [["abhi02@gmail.com"]],
            "isd_code": "91+",
            "Phone": "9212635084",
            "City": [["Delhi"]],
            "Country": [["India"]],
            "Nationality": [["Indian"]],
            "resume_title": [["Software Developer"]],
            "Gender": [["Male"]],
            "notice_period": "15 days",
            # Education done with db
            "qualification_id": [["B.Sc"]],
            "university": [["DU"]],
            "qualification_year": "2011",
            "country_id": [["India"]],
            "major": [["Computer Science"]],
            # Tech Skills done with db
            "total_exp": "5 years",
            "relevent_exp": "4 years",
            "primary_communication_language_id": [["English"]],
            "primary_can_write": [["yes"]],
            "primary_can_speak": [["yes"]],
            "primary_lang_profiency": [["No"]],
            "secondary_communication_language_id": [["chinese"]],
            "secondary_can_write": [["yes"]],
            "secondary_can_speak": [["yes"]],
            "secondary_lang_profiency": [["Punjabi"]],
            "skill_id": [["C,AWS,CI/CD"]],
            "Experience": "5 years",
            "Rating": "7/10 Rating",
            # Company Info done with db
            "employer_name": [["Nityo"]],
            "salary": "45000 INR",
            "Currency": [["INR"]],
            "designation": [["Software Engineer"]],
            "joining_month": [["Aug"]],
            "join_yr": "2018",
            "end_month": [["july"]],
            "end_Year": "2021",
            "role": [["Lead"]],
            "job_city": [["delhi"]],
            "job_country": [["India"]],
            # job looking for done with db
            "work_auth_country": [["India"]],
            "preferred_country": [["US,UK"]],
            "expected_job_type": [["Permanent"]],
            "expected_salary": "100000 INR ",
            "expected_salary_currency": [["INR"]],
            #  27 New Labels
            "Passport": "319985",
            "country_id": [["India"]],
            "dob": "16-04-1998",
            "marital_status": [["Single"]],
            "temp_pin": "110031",
            "temp_add": [["a-76 karkardooma east zone"]],
            "permanent_pin": "110006",
            "permanent_add": [["D-6 , Chandni chowk Delhi -06"]],
            "hometown": [["Delhi"]],
            "marks_grade": [["CGPA 6.7"]],
            "edu_full_part": [["Full time"]],
            "department": [["Tech Team"]],
            "project_name": [["Resume Parser"]],
            "team_size": "4",
            "project_timeline": "6 months",
            "rol_res": [["DEsigning and development"]],
            "project_start_date": "1-03-2021",
            "project_end_date": "7-06-2021",
            "project_achievement": [["Developed in house deepLearning parser"]],
            "project_summary": [["Automate your Task with Resume or Resume Parsing"]],
            "project_employer": [["Nityo Infotech"]],
            "skill_last_used": [["Recently"]],
            "soft_skills": [["English Speaking,Team Handling ,Project Management"]],
            "work_permit": [["Need H1B Visa"]],
            "skill_version": "2.0 latest",
            "expected_job_full_part": [["Full Time"]],
            "social_links": [["abhi@linked.com,abhi@github.com"]],
        }

        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        # merge_dict_output = blank_dict.copy()
        manual_dict_output = blank_dict.copy()
        for key, value in blank_dict.items():
            blank_dict[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################

        ##################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if manual_dict_output["isd_code"] is None:
            pass
        else:
            manual_dict_output["isd_code"] = "".join(
                str(e) for e in manual_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["isd_code"]))):
                isd_code += i

            manual_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if manual_dict_output["Phone"] is None:
            pass
        else:
            manual_dict_output["Phone"] = "".join(
                str(e) for e in manual_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["Phone"]))):
                Phone += i

            manual_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if manual_dict_output["notice_period"] is None:
            pass
        else:
            manual_dict_output["notice_period"] = "".join(
                str(e) for e in manual_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["notice_period"]))
            ):
                notice_period += i

            manual_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if manual_dict_output["qualification_year"] is None:
            pass
        else:
            manual_dict_output["qualification_year"] = "".join(
                str(e) for e in manual_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["qualification_year"]))
            ):
                qualification_year += i

            manual_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if manual_dict_output["total_exp"] is None:
            pass
        else:
            manual_dict_output["total_exp"] = "".join(
                str(e) for e in manual_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["total_exp"]))
            ):
                total_exp += i

            manual_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if manual_dict_output["relevent_exp"] is None:
            pass
        else:
            manual_dict_output["relevent_exp"] = "".join(
                str(e) for e in manual_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            manual_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if manual_dict_output["Experience"] is None:
            pass
        else:
            manual_dict_output["Experience"] = "".join(
                str(e) for e in manual_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["Experience"]))
            ):
                Experience += i

            manual_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if manual_dict_output["Rating"] is None:
            pass
        else:
            manual_dict_output["Rating"] = "".join(
                str(e) for e in manual_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["Rating"]))):
                Rating += i

            manual_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if manual_dict_output["salary"] is None:
            pass
        else:
            manual_dict_output["salary"] = "".join(
                str(e) for e in manual_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["salary"]))):
                salary += i

            manual_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if manual_dict_output["join_yr"] is None:
            pass
        else:
            manual_dict_output["join_yr"] = "".join(
                str(e) for e in manual_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["join_yr"]))):
                join_yr += i

            manual_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if manual_dict_output["expected_salary"] is None:
            pass
        else:
            manual_dict_output["expected_salary"] = "".join(
                str(e) for e in manual_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["expected_salary"]))
            ):
                expected_salary += i

            manual_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if manual_dict_output["Passport"] is None:
            pass
        else:
            manual_dict_output["Passport"] = "".join(
                str(e) for e in manual_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["Passport"]))):
                Passport += i

            manual_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if manual_dict_output["dob"] is None:
            pass
        else:
            manual_dict_output["dob"] = "".join(
                str(e) for e in manual_dict_output["dob"]
            )
            dob = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["dob"]))):
                dob += i

            manual_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if manual_dict_output["temp_pin"] is None:
            pass
        else:
            manual_dict_output["temp_pin"] = "".join(
                str(e) for e in manual_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", manual_dict_output["temp_pin"]))):
                temp_pin += i

            manual_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if manual_dict_output["permanent_pin"] is None:
            pass
        else:
            manual_dict_output["permanent_pin"] = "".join(
                str(e) for e in manual_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            manual_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if manual_dict_output["team_size"] is None:
            pass
        else:
            manual_dict_output["team_size"] = "".join(
                str(e) for e in manual_dict_output["team_size"]
            )
            team_size = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["team_size"]))
            ):
                team_size += i

            manual_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if manual_dict_output["project_timeline"] is None:
            pass
        else:
            manual_dict_output["project_timeline"] = "".join(
                str(e) for e in manual_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["project_timeline"]))
            ):
                project_timeline += i

            manual_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if manual_dict_output["project_start_date"] is None:
            pass
        else:
            manual_dict_output["project_start_date"] = "".join(
                str(e) for e in manual_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["project_start_date"]))
            ):
                project_start_date += i

            manual_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if manual_dict_output["project_end_date"] is None:
            pass
        else:
            manual_dict_output["project_end_date"] = "".join(
                str(e) for e in manual_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["project_end_date"]))
            ):
                project_end_date += i

            manual_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if manual_dict_output["skill_version"] is None:
            pass
        else:
            manual_dict_output["skill_version"] = "".join(
                str(e) for e in manual_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", manual_dict_output["skill_version"]))
            ):
                skill_version += i

            manual_dict_output["skill_version"] = skill_version

        ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->
        # return manual_dict_output
        # sampleJson = jsonpickle.encode(merge_dict_output)
        sampleJson = jsonpickle.encode(manual_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except ValueError:
        error_status_code_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_image

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
    #     ("Resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
    #     ("Resume_Parser Model Accuracy:", df_json),
    # )
    ###########################################################################################################


###################################################### PDF (.pdf and .pdf with Image) ########################################################


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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
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
        #     ("Resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
        #     ("Resume_Parser Model Accuracy:", df_json),
        # )
        ###########################################################################################################

    except Exception as e:
        error_status_code_normal_pdf = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_normal_pdf


################################################# .pdf with image ###################################################################
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except:
        error_status_code_pdf_with_image = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_pdf_with_image


############################################### Excel and CSV Fn ##########################################################################

################################################ .csv #######################################################################################
def extract_text_from_csv(csv_path):
    # return "Sandeep"
    try:
        start = time.time()
        csv_path = base64.b64decode(csv_path.encode("UTF-8"))
        bio_ = io.BytesIO(csv_path)
        # text_x = pd.read_csv(bio_)
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################
        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "JD_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except:
        error_status_code_csv = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_csv


######################################################### .xlxs #################################################################

# Excel to text -->
def extract_text_from_excel(excel_path):
    # return "hello"
    start = time.time()
    try:
        excel_path = base64.b64decode(excel_path.encode("UTF-8"))
        bio_ = io.BytesIO(excel_path)
        text_x = pd.read_excel(bio_, index_col=0)
        # return text_x
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
        }
        # print(blank_dict)

        ############################ XXXXXXXXXXXXXXXX ##############################

        merge_dict_output = blank_dict.copy()
        for key, value in _df_.items():
            merge_dict_output[key] = value

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        ##################################################

        # Extracting only Numeric values from the dict(values) e.g
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version
        ###################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return loaded_r
    except:
        error_status_code_excel = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_excel


############################################################## Text Fn #################################################################

############################################################### .txt ###################################################################
def extract_text_from_text(text_path):
    # return "Sandeep"
    start = time.time()
    try:
        text_path = base64.b64decode(text_path.encode("UTF-8"))
        bio_ = io.BytesIO(text_path)
        # return bio_
        text_x = pd.read_csv(
            bio_,
            error_bad_lines=False,
            skip_blank_lines=True,
            delim_whitespace=False,
            low_memory=True,
            float_precision=False,
        )

        # text_path = base64.b64decode(text_path.encode("UTF-8"))
        # bio_ = io.BytesIO(text_path)
        # # return bio_
        # # text_x = pd.read_csv(bio_, error_bad_lines=False)
        # # return text_x
        # # print(text_x)
        # text_x = pd.read_csv(bio_)
        # # print(text_x)
        # return text_x
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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
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
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_tmline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version
        # ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
            "data": decodedSet,
        }
        r = json.dumps(manual_text_add, indent=10)
        loaded_r = json.loads(r)
        return text_x
        # return loaded_r
    except:
        error_status_code_text = {
            "status": "failure",
            "status_code": status.HTTP_500_INTERNAL_SERVER_ERROR,
            "message": "Internal Server Error",
        }
    return error_status_code_text


########################################################## .rtf  ############################################################################


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
                # Personal Details
                "FN": "first_name",
                "LN": "last_name",
                "EM": "Email",
                "ISD": "isd_code",
                "PH": "Phone",  # Change
                "CT": "City",
                "LOC": "Country",
                "NAT": "Nationality",
                "TTL": "resume_title",
                "GEN": "Gender",
                "NP": "notice_period",
                # Education
                "HQL": "qualification_id",
                "UNV": "university",
                "PYR": "qualification_year",
                "ECTR": "country_id",
                "SPL": "major",
                # Tech Skills
                "TE": "total_exp",
                "RE": "relevent_exp",
                "PLG": "primary_communication_language_id",  # Change
                "PLW": "primary_can_write",  # Change
                "PLS": "primary_can_speak",  # Change
                "PLP": "primary_lang_profiency",  # Change
                "SLG": "secondary_communication_language_id",  # Change
                "SLW": "secondary_can_write",  # Change
                "SLS": "secondary_can_speak",  # Change
                "SLP": "secondary_lang_profiency",  # Change
                "TECS": "skill_id",
                "EXPS": "Experience",
                "RAT": "Rating",
                # Company Info
                "COM": "employer_name",
                "SAL": "salary",
                "CUR": "Currency",
                "DES": "designation",
                "J_MN": "joining_month",
                "J_YR": "join_yr",  # without db field
                "E_MN": "end_month",
                "E_YR": "end_Year",
                "ROL": "role",
                "J_CT": "job_city",  # Change
                "J_CTR": "job_country",  # Change
                # job looking for done with db
                "WAC": "work_auth_country",
                "PJC": "preferred_country",
                "EJP_EJC": "expected_job_permanent_contract",  # Change
                "ESAL": "expected_salary",
                "ESALC": "expected_salary_currency",
                # Labels
                "PN": "Passport",  # done with db field name
                "CTY": "country_id",  # done with db field name
                "DOB": "dob",  # done with db field name
                "MAR": "marital_status",  # done with db field name
                "TPIN": "temp_pin",
                "TA": "temp_add",
                "PPIN": "permanent_pin",  # Change
                "PA": "permanent_add",  # Change
                "HTW": "Hometown",
                "M_G": "marks_grade",  # Change
                "EF_EP": "edu_full_part",  # Change
                "DEP": "Department",
                "PN": "project_name",
                "TS": "team_size",
                "PT": "project_timeline",  # Change
                "RESP": "rol_res",
                "PSD": "project_start_date",  # Change
                "PED": "project_end_date",  # Change
                "PACH": "project_achievement",  # Change
                "PS": "project_summary",
                "PC": "project_employer",
                "LUSE": "skill_last_used",  # Change
                "SFTS": "soft_skills",
                "WP": "work_permit",
                "VER": "skill_version",  # Change
                "EJF_EJP": "expected_job_full_part",  # Change
                "SLNK": "social_links",
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
            "first_name": None,
            "last_name": None,
            "Email": None,
            "isd_code": None,
            "Phone": None,
            "City": None,
            "Country": None,
            "Nationality": None,
            "resume_title": None,
            "Gender": None,
            "notice_period": None,
            # Education done with db
            "qualification_id": None,
            "university": None,
            "qualification_year": None,
            "country_id": None,
            "major": None,
            # Tech Skills done with db
            "total_exp": None,
            "relevent_exp": None,
            "primary_communication_language_id": None,
            "primary_can_write": None,
            "primary_can_speak": None,
            "primary_lang_profiency": None,
            "secondary_communication_language_id": None,
            "secondary_can_write": None,
            "secondary_can_speak": None,
            "secondary_lang_profiency": None,
            "skill_id": None,
            "Experience": None,
            "Rating": None,
            # Company Info done with db
            "employer_name": None,
            "salary": None,
            "currency": None,
            "designation": None,
            "joining_month": None,
            "join_yr": None,
            "end_month": None,
            "end_Year": None,
            "role": None,
            "job_city": None,
            "job_country": None,
            # job looking for done with db
            "work_auth_country": None,
            "preferred_country": None,
            "expected_job_type": None,
            "expected_salary": None,
            "expected_salary_currency": None,
            # Labels
            "Passport": None,
            "country_id": None,
            "dob": None,
            "marital_status": None,
            "temp_pin": None,
            "temp_add": None,
            "permanent_pin": None,
            "permanent_add": None,
            "hometown": None,
            "marks_grade": None,
            "edu_full_part": None,
            "department": None,
            "project_name": None,
            "team_size": None,
            "project_timeline": None,
            "rol_res": None,
            "project_start_date": None,
            "project_end_date": None,
            "project_achievement": None,
            "project_summary": None,
            "project_employer": None,
            "skill_last_used": None,
            "soft_skills": None,
            "work_permit": None,
            "skill_version": None,
            "expected_job_full_part": None,
            "social_links": None,
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
        # 1) isd_code, 2) Phone, 3) notice_period, 4) qualification_year, 5) total_exp, 6) relevent_exp, 7) Experience, 8) Rating, 9) salary, 10) join_yr  11)expected_salary
        #  12)Passport 13)dob 14)temp_pin 15)prmnt_pin 16) team_size 17) project_tmline 18) start_date 19)end_date 20) version

        # 1) for "isd_code"  -->
        if merge_dict_output["isd_code"] is None:
            pass
        else:
            merge_dict_output["isd_code"] = "".join(
                str(e) for e in merge_dict_output["isd_code"]
            )
            isd_code = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["isd_code"]))):
                isd_code += i

            merge_dict_output["isd_code"] = isd_code

        # # 2) for "Phone"  -->
        if merge_dict_output["Phone"] is None:
            pass
        else:
            merge_dict_output["Phone"] = "".join(
                str(e) for e in merge_dict_output["Phone"]
            )
            Phone = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Phone"]))):
                Phone += i

            merge_dict_output["Phone"] = Phone

        # 3) for "notice_period"  -->
        if merge_dict_output["notice_period"] is None:
            pass
        else:
            merge_dict_output["notice_period"] = "".join(
                str(e) for e in merge_dict_output["notice_period"]
            )
            notice_period = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["notice_period"]))
            ):
                notice_period += i

            merge_dict_output["notice_period"] = notice_period

        # 4) for "qualification_year"  -->
        if merge_dict_output["qualification_year"] is None:
            pass
        else:
            merge_dict_output["qualification_year"] = "".join(
                str(e) for e in merge_dict_output["qualification_year"]
            )
            qualification_year = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["qualification_year"]))
            ):
                qualification_year += i

            merge_dict_output["qualification_year"] = qualification_year

        # 5) for "total_exp"  -->
        if merge_dict_output["total_exp"] is None:
            pass
        else:
            merge_dict_output["total_exp"] = "".join(
                str(e) for e in merge_dict_output["total_exp"]
            )
            total_exp = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["total_exp"]))):
                total_exp += i

            merge_dict_output["total_exp"] = total_exp

        # 6) for "relevent_exp"  -->
        if merge_dict_output["relevent_exp"] is None:
            pass
        else:
            merge_dict_output["relevent_exp"] = "".join(
                str(e) for e in merge_dict_output["relevent_exp"]
            )
            relevent_exp = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["relevent_exp"]))
            ):
                relevent_exp += i

            merge_dict_output["relevent_exp"] = relevent_exp

        # 7) for "Experience"  -->
        if merge_dict_output["Experience"] is None:
            pass
        else:
            merge_dict_output["Experience"] = "".join(
                str(e) for e in merge_dict_output["Experience"]
            )
            Experience = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["Experience"]))
            ):
                Experience += i

            merge_dict_output["Experience"] = Experience

        # 8) for "Rating"  -->
        if merge_dict_output["Rating"] is None:
            pass
        else:
            merge_dict_output["Rating"] = "".join(
                str(e) for e in merge_dict_output["Rating"]
            )
            Rating = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Rating"]))):
                Rating += i

            merge_dict_output["Rating"] = Rating

        # 9) for "salary"  -->
        if merge_dict_output["salary"] is None:
            pass
        else:
            merge_dict_output["salary"] = "".join(
                str(e) for e in merge_dict_output["salary"]
            )
            salary = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["salary"]))):
                salary += i

            merge_dict_output["salary"] = salary

        # 10) for "join_yr"  -->
        if merge_dict_output["join_yr"] is None:
            pass
        else:
            merge_dict_output["join_yr"] = "".join(
                str(e) for e in merge_dict_output["join_yr"]
            )
            join_yr = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["join_yr"]))):
                join_yr += i

            merge_dict_output["join_yr"] = join_yr

        # 11) for "expected_salary"  -->
        if merge_dict_output["expected_salary"] is None:
            pass
        else:
            merge_dict_output["expected_salary"] = "".join(
                str(e) for e in merge_dict_output["expected_salary"]
            )
            expected_salary = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["expected_salary"]))
            ):
                expected_salary += i

            merge_dict_output["expected_salary"] = expected_salary

        # # 12) for "Passport "  -->
        if merge_dict_output["Passport"] is None:
            pass
        else:
            merge_dict_output["Passport"] = "".join(
                str(e) for e in merge_dict_output["Passport"]
            )
            Passport = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["Passport"]))):
                Passport += i

            merge_dict_output["Passport"] = Passport

        # 13) for "dob"  -->
        if merge_dict_output["dob"] is None:
            pass
        else:
            merge_dict_output["dob"] = "".join(str(e) for e in merge_dict_output["dob"])
            dob = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["dob"]))):
                dob += i

            merge_dict_output["dob"] = dob

        # 14) for "temp_pin"  -->
        if merge_dict_output["temp_pin"] is None:
            pass
        else:
            merge_dict_output["temp_pin"] = "".join(
                str(e) for e in merge_dict_output["temp_pin"]
            )
            temp_pin = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["temp_pin"]))):
                temp_pin += i

            merge_dict_output["temp_pin"] = temp_pin

        # 15) for "permanent_pin"  -->
        if merge_dict_output["permanent_pin"] is None:
            pass
        else:
            merge_dict_output["permanent_pin"] = "".join(
                str(e) for e in merge_dict_output["permanent_pin"]
            )
            permanent_pin = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["permanent_pin"]))
            ):
                permanent_pin += i

            merge_dict_output["permanent_pin"] = permanent_pin

        # 16) for "team_size"  -->
        if merge_dict_output["team_size"] is None:
            pass
        else:
            merge_dict_output["team_size"] = "".join(
                str(e) for e in merge_dict_output["team_size"]
            )
            team_size = int()
            for i in list(map(int, re.findall(r"\d+", merge_dict_output["team_size"]))):
                team_size += i

            merge_dict_output["team_size"] = team_size

        # 17) for "project_timeline"  -->
        if merge_dict_output["project_timeline"] is None:
            pass
        else:
            merge_dict_output["project_timeline"] = "".join(
                str(e) for e in merge_dict_output["project_timeline"]
            )
            project_timeline = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_timeline"]))
            ):
                project_timeline += i

            merge_dict_output["project_timeline"] = project_timeline

        # 18) for "project_start_date"  -->
        if merge_dict_output["project_start_date"] is None:
            pass
        else:
            merge_dict_output["project_start_date"] = "".join(
                str(e) for e in merge_dict_output["project_start_date"]
            )
            project_start_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_start_date"]))
            ):
                project_start_date += i

            merge_dict_output["project_start_date"] = project_start_date

        # 19) for "project_end_date"  -->
        if merge_dict_output["project_end_date"] is None:
            pass
        else:
            merge_dict_output["project_end_date"] = "".join(
                str(e) for e in merge_dict_output["project_end_date"]
            )
            project_end_date = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["project_end_date"]))
            ):
                project_end_date += i

            merge_dict_output["project_end_date"] = project_end_date

        # 20) for "skill_version"  -->
        if merge_dict_output["skill_version"] is None:
            pass
        else:
            merge_dict_output["skill_version"] = "".join(
                str(e) for e in merge_dict_output["skill_version"]
            )
            skill_version = int()
            for i in list(
                map(int, re.findall(r"\d+", merge_dict_output["skill_version"]))
            ):
                skill_version += i

            merge_dict_output["skill_version"] = skill_version

        ####################################### XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ####################################

        # Converting dict to json by serialization -->

        sampleJson = jsonpickle.encode(merge_dict_output)
        decodedSet = jsonpickle.decode(sampleJson)
        manual_text_add = {
            "status": "success",
            "status_code": status.HTTP_200_OK,
            "message": "Resume file Uploaded Successfully",
            # "Resume_Parser NLP Model running time in seconds": round(script_time, 1),
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


############################################## Process Fn #############################################################################
def process(file_content, file_extension):
    """
    Wrapper function to detect the file extension and call text
    extraction function accordingly

    :param file_path: path of file of which text is to be extracted
    :param extension: extension of file `file_name`
    """
    Resume_lines = ""
    if file_extension == ".docx":
        Resume_lines = extract_text_from_docx(file_content)
        return Resume_lines
    elif file_extension == ".DOCX":
        Resume_lines = extract_text_from_docx(file_content)
        return Resume_lines
    elif file_extension == ".doc":
        Resume_lines = extract_text_from_doc(file_content)
        return Resume_lines
    elif file_extension == ".DOC":
        Resume_lines = extract_text_from_doc(file_content)
        return Resume_lines
    # elif file_extension == ".doc":
    #     Resume_lines = extract_text_from_doc_with_image(file_content)
    #     return Resume_lines
    # elif file_extension == ".DOC":
    #     Resume_lines = extract_text_from_doc_with_image(file_content)
    #     return Resume_lines
    elif file_extension == ".pdf":
        Resume_lines = extract_text_from_pdf(file_content)
        return Resume_lines
    elif file_extension == ".PDF":
        Resume_lines = extract_text_from_pdf(file_content)
        return Resume_lines
    elif file_extension == ".pdf":
        Resume_lines = extract_text_from_pdf_with_image(file_content)
        return Resume_lines
    elif file_extension == ".PDF":
        Resume_lines = extract_text_from_pdf_with_image(file_content)
        return Resume_lines
    elif file_extension == ".png":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".PNG":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".jpg":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".JPG":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".jpeg":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".JPEG":
        Resume_lines = extract_text_from_image(file_content)
        return Resume_lines
    elif file_extension == ".csv":
        Resume_lines = extract_text_from_csv(file_content)
        return Resume_lines
    elif file_extension == ".CSV":
        Resume_lines = extract_text_from_csv(file_content)
        return Resume_lines
    elif file_extension == ".xlxs":
        Resume_lines = extract_text_from_excel(file_content)
        return Resume_lines
    elif file_extension == ".XLXS":
        Resume_lines = extract_text_from_excel(file_content)
        return Resume_lines
    elif file_extension == ".txt":
        Resume_lines = extract_text_from_text(file_content)
        return Resume_lines
    elif file_extension == ".TXT":
        Resume_lines = extract_text_from_text(file_content)
        return Resume_lines
    elif file_extension == ".rtf":
        Resume_lines = extract_text_from_rtf(file_content)
        return Resume_lines
    elif file_extension == ".RTF":
        Resume_lines = extract_text_from_rtf(file_content)
        return Resume_lines
    else:
        error_status_code = {
            "status": "failure",
            "status_code": status.HTTP_422_UNPROCESSABLE_ENTITY,
            "message": "file_content and file_extension Mismatch/Missing",
        }
        return error_status_code


########################################## XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX ############################################