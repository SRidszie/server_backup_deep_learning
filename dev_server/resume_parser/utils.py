import time

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


# DF conversion of output
import pandas as pd
from dframcy import DframCy



# Set URL
url = "http://127.0.0.1:8000/rs/resume-parse/" # for local
###########################################################


# url_for_domain = "https://devparse.mycareercube.com/jd-parse/"


nlp = spacy.load("en_core_web_sm")
dframcy = DframCy(nlp)


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
            "TTL": "Title",
            "JD_TY": "JD_TYPE",
            "JD_PRTY": "JD_PRIORITY",
            "JD_CAT": "JD_CATEGORY",
            "CLT_TY": "CLIENT_TYPE",
            "ENT_NAM": "Entity_ID",
            "EMP_TY": "Employment_Type",
            "Job_LOC_C": "workzone_id",
            "DEP": "Department_Functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "Num_of_certify_cv",
            "TOT_CV": "Num_of_cv",
            "VAC": "Jd_openings",
            "EXP_QUAL": "Expertise_qualification",
            "ROL_RES": "Roles_responsibility",
            "SAL_MN": "Salary_min",
            "SAL_MX": "Salary_max",
            "CLT_NM": "client_id",
            "EDU": "Educational_Qualification",
            "CNT_PRD": "Contact_Period",
            "SFT_SK": "SOFT_SKILLS",
            "GEN": "GENDER",
            "CTRY": "COUNTRY",
            "JON_DT": "JOINING_DATE",
            "SHFT_TM": "SHIFT_TIMINGS",
            "PH": "PHONE",
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(set)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    ## Adding Accuracy Matrix in API:

    # Opening JSON file
    f = open("metrics.json", "r")

    # returns JSON object as
    # a dictionary
    df_json = json.loads(f.read())

    # print(df_json)

    ## Adding Status Code in the API

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    return (
        _df_,
        ("resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
        ("resume_Parser Model Accuracy:", df_json),
    )


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
            "TTL": "Title",
            "JD_TY": "JD_TYPE",
            "JD_PRTY": "JD_PRIORITY",
            "JD_CAT": "JD_CATEGORY",
            "CLT_TY": "CLIENT_TYPE",
            "ENT_NAM": "Entity_ID",
            "EMP_TY": "Employment_Type",
            "Job_LOC_C": "workzone_id",
            "DEP": "Department_Functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "Num_of_certify_cv",
            "TOT_CV": "Num_of_cv",
            "VAC": "Jd_openings",
            "EXP_QUAL": "Expertise_qualification",
            "ROL_RES": "Roles_responsibility",
            "SAL_MN": "Salary_min",
            "SAL_MX": "Salary_max",
            "CLT_NM": "client_id",
            "EDU": "Educational_Qualification",
            "CNT_PRD": "Contact_Period",
            "SFT_SK": "SOFT_SKILLS",
            "GEN": "GENDER",
            "CTRY": "COUNTRY",
            "JON_DT": "JOINING_DATE",
            "SHFT_TM": "SHIFT_TIMINGS",
            "PH": "PHONE",
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(set)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    ## Adding Accuracy Matrix in API:

    # Opening JSON file
    f = open("metrics.json", "r")

    # returns JSON object as
    # a dictionary
    df_json = json.loads(f.read())

    ## Adding Status Code in the API

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    return (
        _df_,
        ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
        ("JD_Parser Model Accuracy:", df_json),
    )


def extract_text_from_image(image_file):
    start = time.time()
    image = Image.open(image_file)
    try:
        image.filter(ImageFilter.SHARPEN)
    except ValueError:
        print("Got an image that failed to sharpen", image_file)
        pass
    text = pytesseract.image_to_string(image)
    text = text.replace("\n", " ")

    ## Spacy - NLP

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(text)
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )

    ## Rearrange Column Names
    df = entity_text_dataframe[["ent_label", "ent_text"]]

    ## Renaming Columns as per business need:

    df["ent_label"] = df["ent_label"].replace(
        {
            "TTL": "Title",
            "JD_TY": "JD_TYPE",
            "JD_PRTY": "JD_PRIORITY",
            "JD_CAT": "JD_CATEGORY",
            "CLT_TY": "CLIENT_TYPE",
            "ENT_NAM": "Entity_ID",
            "EMP_TY": "Employment_Type",
            "Job_LOC_C": "workzone_id",
            "DEP": "Department_Functional_area_id",
            "MAN_SK": "key_skills_and",
            "OPT_SK": "key_skills_or",
            "SK_MT": "description",
            "CRT_CV": "Num_of_certify_cv",
            "TOT_CV": "Num_of_cv",
            "VAC": "Jd_openings",
            "EXP_QUAL": "Expertise_qualification",
            "ROL_RES": "Roles_responsibility",
            "SAL_MN": "Salary_min",
            "SAL_MX": "Salary_max",
            "CLT_NM": "client_id",
            "EDU": "Educational_Qualification",
            "CNT_PRD": "Contact_Period",
            "SFT_SK": "SOFT_SKILLS",
            "GEN": "GENDER",
            "CTRY": "COUNTRY",
            "JON_DT": "JOINING_DATE",
            "SHFT_TM": "SHIFT_TIMINGS",
            "PH": "PHONE",
        }
    )

    df_ = (
        df.groupby("ent_label", sort=True)["ent_text"]
        .apply(set)
        .reset_index(name="ent_text")
    )

    ## dataframe to dict

    _df_ = df_.set_index("ent_label").T.to_dict("list")

    ## Adding Accuracy Matrix in API:

    import json

    # Opening JSON file
    f = open("metrics.json", "r")

    # returns JSON object as
    # a dictionary
    df_json = json.loads(f.read())

    ## Adding Status Code in the API

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    return (
        _df_,
        ("resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
        ("resume_Parser Model Accuracy:", df_json),
    )


def extract_text_from_image_path(image_file_path):
    with open(image_file_path, "rb") as image_file:
        return extract_text_from_image(image_file)



def extract_text_from_docx(doc_path):
    start = time.time()
    """
    Helper function to extract plain text from .docx files

    :param doc_path: path to .docx file to be extracted
    :return: string of extracted text
    
    """

    try:
        temp = docx2txt.process(doc_path)
        text = [line.replace("\t", " ") for line in temp.split("\n") if line]

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
                "TTL": "Title",
                "JD_TY": "JD_TYPE",
                "JD_PRTY": "JD_PRIORITY",
                "JD_CAT": "JD_CATEGORY",
                "CLT_TY": "CLIENT_TYPE",
                "ENT_NAM": "Entity_ID",
                "EMP_TY": "Employment_Type",
                "Job_LOC_C": "workzone_id",
                "DEP": "Department_Functional_area_id",
                "MAN_SK": "key_skills_and",
                "OPT_SK": "key_skills_or",
                "SK_MT": "description",
                "CRT_CV": "Num_of_certify_cv",
                "TOT_CV": "Num_of_cv",
                "VAC": "Jd_openings",
                "EXP_QUAL": "Expertise_qualification",
                "ROL_RES": "Roles_responsibility",
                "SAL_MN": "Salary_min",
                "SAL_MX": "Salary_max",
                "CLT_NM": "client_id",
                "EDU": "Educational_Qualification",
                "CNT_PRD": "Contact_Period",
                "SFT_SK": "SOFT_SKILLS",
                "GEN": "GENDER",
                "CTRY": "COUNTRY",
                "JON_DT": "JOINING_DATE",
                "SHFT_TM": "SHIFT_TIMINGS",
                "PH": "PHONE",
            }
        )

        df_ = (
            df.groupby("ent_label", sort=True)["ent_text"]
            .apply(set)
            .reset_index(name="ent_text")
        )

        ## dataframe to dict

        _df_ = df_.set_index("ent_label").T.to_dict("list")

        ## Adding Accuracy Matrix in API:

        # Opening JSON file
        f = open("metrics.json", "r")

        # returns JSON object as
        # a dictionary
        df_json = json.loads(f.read())

    
        ## Adding Status Code in the API

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        return (
            _df_,
            ("resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
            ("reusme_Parser Model Accuracy:", df_json),
        )

      
    except KeyError:
        return " "



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
                "TTL": "Title",
                "JD_TY": "JD_TYPE",
                "JD_PRTY": "JD_PRIORITY",
                "JD_CAT": "JD_CATEGORY",
                "CLT_TY": "CLIENT_TYPE",
                "ENT_NAM": "Entity_ID",
                "EMP_TY": "Employment_Type",
                "Job_LOC_C": "workzone_id",
                "DEP": "Department_Functional_area_id",
                "MAN_SK": "key_skills_and",
                "OPT_SK": "key_skills_or",
                "SK_MT": "description",
                "CRT_CV": "Num_of_certify_cv",
                "TOT_CV": "Num_of_cv",
                "VAC": "Jd_openings",
                "EXP_QUAL": "Expertise_qualification",
                "ROL_RES": "Roles_responsibility",
                "SAL_MN": "Salary_min",
                "SAL_MX": "Salary_max",
                "CLT_NM": "client_id",
                "EDU": "Educational_Qualification",
                "CNT_PRD": "Contact_Period",
                "SFT_SK": "SOFT_SKILLS",
                "GEN": "GENDER",
                "CTRY": "COUNTRY",
                "JON_DT": "JOINING_DATE",
                "SHFT_TM": "SHIFT_TIMINGS",
                "PH": "PHONE",
            }
        )

     
        df_ = (
            df.groupby("ent_label", sort=True)["ent_text"]
            .apply(set)
            .reset_index(name="ent_text")
        )

        ## dataframe to dict

        _df_ = df_.set_index("ent_label").T.to_dict("list")

       
        ## Adding Accuracy in API:

        # Opening JSON file
        f = open("metrics.json", "r")

        # returns JSON object as
        # a dictionary
        df_json = json.loads(f.read())

         ## Adding Status Code in the API

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        return (
            _df_,
            ("resume_Parser NLP Model running time in seconds:", round(script_time, 1)),
            ("resume_Parser Model Accuracy:", df_json),
        )
  

    except Exception as e:
        logging.error("Error in pdf file:: " + str(e))
        return []

def process(file):
    # import pdb
    # pdb.set_trace()
    """
    Main function to process resume file to json.
    :param file: Resume file
    :return: resume_data: Parsed resume dictionary
    """
    if file.name.endswith("docx"):
        resume_lines = extract_text_from_docx(file)
        # print(resume_lines)
        return resume_lines
    # elif file.name.endswith("doc"):
    #     resume_lines = obtaintext(file)
    #     # print(resume_lines)
    #     return resume_lines
    elif file.name.endswith("PNG"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("png"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("JPG"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("jpg"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("jpeg"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("JPEG"):
        resume_lines = extract_text_from_image(file)
        # print(resume_lines)
        return resume_lines
    elif file.name.endswith("pdf"):
        resume_lines = extract_text_from_pdf(file)
        # print(resume_lines)
        return resume_lines
    else:
        return None

