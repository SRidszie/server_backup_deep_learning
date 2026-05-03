import time

# import MySQLdb
import pytesseract
from PIL import Image, ImageFilter

# import fitz
# from PIL import ImageFilter
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

from shutil import copyfile

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

import mysql.connector

# DF conversion of output
import pandas as pd
from dframcy import DframCy

# mydb = mysql.connector.connect(
#     host="localhost", username="root", password="root", database="dev_ai"
# )

# mycursor = mydb.cursor()

# Set URL
url = "http://127.0.0.1:8000/api/"

# nlp = spacy.load("en_core_web_sm")
nlp = spacy.load("en_jd_parser")
dframcy = DframCy(nlp)


def extract_text_from_doc(doc_path):
    start = time.time()
    temp = textract.process(str(doc_path)).decode("utf-8")
    text = [line.replace("\t", " ") for line in temp.split("\n") if line]

    # return text
    # return " ".join(text)

    # Spacy - NLP

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(str(text))
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )
    # print(entity_text_dataframe)

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

    # print(df_json)

    ## Adding Status Code in the API

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    return (
        _df_,
        ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
        ("JD_Parser Model Accuracy:", df_json),
    )


def extract_text_from_pdf_with_image(pdf_file):
    start = time.time()
    final_text = []
    images = pdf2image.convert_from_path(str(pdf_file))
    for pages, img in enumerate(images):
        text = pytesseract.image_to_string(img)
        # print(pages,text)
        # return (text)
        final_text.append({pages, text})
        # print(final_text)
        # doc = nlp(text)
        # text_ = [entity.text for entity in doc.ents]
        # label = [entity.label_ for entity in doc.ents]
        # c = dict(zip(text_, label))
        # print(c)
        # print(pages,text)
        # return(c)
        # return(pages,(text))

    # return final_text

    ## Spacy - NLP

    ## Converting Spacy output into pandas dataframe
    text_df = dframcy.nlp(str(final_text))
    token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
        text_df, separate_entity_dframe=True
    )
    # print(entity_text_dataframe)

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

    # print(df_json)

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
    # print(entity_text_dataframe)

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

    # print(df_json)

    ## Adding Status Code in the API

    ## For time Calculation - Run this at last
    script_time = time.time() - start
    return (
        _df_,
        ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
        ("JD_Parser Model Accuracy:", df_json),
    )


def extract_text_from_image_path(image_file_path):
    with open(image_file_path, "rb") as image_file:
        return extract_text_from_image(image_file)


def extract_text_from_image_url(image_url):
    image_data = requests.get(image_url, stream=True).raw
    return extract_text_from_image(image_data)


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
        # print(entity_text_dataframe)

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

        # print(df_json)

        ## Adding Status Code in the API

        ## For time Calculation - Run this at last
        script_time = time.time() - start
        return (
            _df_,
            ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
            ("JD_Parser Model Accuracy:", df_json),
        )

        # return ' '.join(text)
    except KeyError:
        return " "


# def extract_text_from_doc(doc_path):
#     '''
#     Helper function to extract plain text from .doc files

#     :param doc_path: path to .doc file to be extracted
#     :return: string of extracted text
#     '''
#     try:
#         try:
#             import textract
#         except ImportError:
#             return ' '
#         temp = textract.process(doc_path).decode('utf-8')
#         text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
#         return ' '.join(text)
#     except KeyError:
#         return ' '


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
        # text_ = [entity.text for entity in doc.ents]
        # label = [entity.label_ for entity in doc.ents]

        ## NLP Output in Dict Format
        # NLP Output in JSON - Format 1
        # c = dict(zip(text_, label))
        # Additional Manual Labels
        # jdtype = ["a", "b"]

        ## NLP Output in DataFrame format - directly using dframcy (spacy add on)
        # for separate entity_text dataframe
        text_df = dframcy.nlp(text)
        token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
            text_df, separate_entity_dframe=True
        )
        # print(entity_text_dataframe)

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

        #################################################################################################

        # df.set_index("ent_label", inplace=True)

        # df_ = df.groupby("ent_label", sort=True)["ent_text"].apply(set)

        ###################################################################################################

        df_ = (
            df.groupby("ent_label", sort=True)["ent_text"]
            .apply(set)
            .reset_index(name="ent_text")
        )

        ## dataframe to dict

        _df_ = df_.set_index("ent_label").T.to_dict("list")

        ## change indexing
        # setting first name as index column
        # df_.set_index("ent_label", inplace=True)

        # df_.to_csv("transpose.csv")
        # df_ = df.groupby("ent_label", sort=True)
        # df_ = (
        #     df.groupby("ent_label", sort=True)["ent_text"]
        #     .apply(set)
        #     .reset_index(name="ent_text")
        # )

        # _df_ = df_.to_dict("index")
        # print(list(df_))
        # print(token_annotation_dataframe)
        # For loop to
        # for i in range(100):
        #     data = json.loads(requests.get(
        #         url=url,
        #         params={'page': c}
        #     ).text)['results']

        # data_norm = pd.read_json(json.dumps(data))
        # print(data_norm)
        # print("JD_Parser NLP Model running time:", round(script_time, 1), "seconds")
        # sql = """ insert into prodgy (name) VALUES  ('" + c + "') """
        # sql = "INSERT INTO prodgy (name) VALUES (%s) "

        # # `dev_ai`.`prodgy` (`name`) VALUES ('nidhi');
        # # sql="insert into dev_ai.prodgy (name) values ('ni')"
        # # mycursor.execute(sql)
        # # print(mycursor.execute(sql))
        # mycursor.execute(sql, (json.dumps(c),))
        # # print(mycursor.execute(sql))
        # mydb.commit()

        # print(mycursor.rowcount, "record inserted.")
        # mycursor.execute("SELECT * FROM dev_ai.prodgy ")
        # results = mycursor.fetchmany(10)

        # print(results)
        # mycursor.close()

        ###############################################################################################################################
        # 1) Method 1 - Directly using pandas SQL

        # import the module
        # from sqlalchemy import create_engine

        # # create sqlalchemy engine
        # engine = create_engine(
        #     "mysql+pymysql://{user}:{pw}@localhost/{db}".format(
        #         user="root", pw="Tuffy@77111", db="spacy_nlp"
        #     )
        # )

        # # Insert whole DataFrame into MySQL
        # df_.to_sql(
        #     name="jd_details_api",
        #     con=engine,
        #     if_exists="append",
        #     index=False,
        #     index_label=None,
        #     chunksize=None,
        #     dtype=None,
        #     method=None,
        # )
        # #################################################################################################################################

        ## Adding Accuracy in API:

        import json

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
            ("JD_Parser NLP Model running time in seconds:", round(script_time, 1)),
            ("JD_Parser Model Accuracy:", df_json),
        )
        # mycursor.close()
        # mycursor.execute("SELECT * FROM dev_ai.prodgy ")
        # results = mycursor.fetch()

        # print(results)
        # mydb = dbconnect()
        # cursor = mydb.cursor()
        # sql = """ INSERT INTO 'dev_ai'.'accounts' ('name') VALUES (%s) """
        # mycursor.execute(sql,(json.dumps(c),))
        # mycursor.commit()
        # print(mycursor.rowcount, "record inserted.")
        # mycursor.close()
        # return text
        # json_data = pymysql.escape_string(c)
        # mycursor.execute("insert into 'dev_ai'.'accounts' ('name') values ('" + c + "'); ")
        # INSERT INTO `dev_ai`.`accounts` (`name`) VALUES ('Riddhi');
        # mycursor.close()

        # print(results)

        # Normalize a bit, removing line breaks
        # full_string = full_string.replace("\r", "\n")
        # full_string = full_string.replace("\t", " ")

        # Remove awkward LaTeX bullet characters
        # full_string = re.sub(r"\(cid:\d{0,2}\)", " ", full_string)

        # Split text blob into individual lines
        # resume_lines = full_string.splitlines(True)

        # Remove empty strings and whitespaces
        # resume_lines = [re.sub('\s+', ' ', line.strip()) for line in resume_lines if line.strip()]

        # return resume_lines

    except Exception as e:
        logging.error("Error in pdf file:: " + str(e))
        return []


# import subprocess


# def extract_text_from_document(doc_path):
#     temp = subprocess.check_output(["antiword", "-m", "utf-8.txt", doc_path])
#     text = [line.replace("\t", " ") for line in temp.split("\n") if line]
#     return " ".join(text)


# import glob, os


# def readfiles(file):
#     # pdf=pdf2image.convert_from_path(file)
#     # files_path=pdf
#     files_path = [os.path.abspath(x) for x in os.listdir()]
#     files_path = [pdf for pdf in files_path if ".pdf" in pdf]
#     pdfs = []
#     for file in files_path:
#         text = textract.process(file, method="tesseract", language="eng")

#         pdfs += [text]


#    os.chdir(path)
#    images = pdf2image.convert_from_path(str(path))
#    pdfs = []
#    for file in glob.glob("*.pdf"):
#        print(file)
#        pdfs.append(file)


## Docx fun 2 - Unused by working
# def obtaintext(docFileName):
#     start = time.time()
#     document = Document(docFileName)
#     finalText = []
#     for line in document.paragraphs:
#         finalText.append(line.text)

#     ## Spacy - NLP

#     ## Converting Spacy output into pandas dataframe
#     text_df = dframcy.nlp(str(finalText))
#     token_annotation_dataframe, entity_text_dataframe = dframcy.to_dataframe(
#         text_df, separate_entity_dframe=True
#     )
#     # print(entity_text_dataframe)

#     ## Rearrange Column Names
#     df = entity_text_dataframe[["ent_label", "ent_text"]]

#     ## Renaming Columns as per business need:

#     df["ent_label"] = df["ent_label"].replace(
#         {
#             "TTL": "Title",
#             "JD_TY": "JD_TYPE",
#             "JD_PRTY": "JD_PRIORITY",
#             "JD_CAT": "JD_CATEGORY",
#             "CLT_TY": "CLIENT_TYPE",
#             "ENT_NAM": "Entity_ID",
#             "EMP_TY": "Employment_Type",
#             "Job_LOC_C": "workzone_id",
#             "DEP": "Department_Functional_area_id",
#             "MAN_SK": "key_skills_and",
#             "OPT_SK": "key_skills_or",
#             "SK_MT": "description",
#             "CRT_CV": "Num_of_certify_cv",
#             "TOT_CV": "Num_of_cv",
#             "VAC": "Jd_openings",
#             "EXP_QUAL": "Expertise_qualification",
#             "ROL_RES": "Roles_responsibility",
#             "SAL_MN": "Salary_min",
#             "SAL_MX": "Salary_max",
#             "CLT_NM": "client_id",
#             "EDU": "Educational_Qualification",
#             "CNT_PRD": "Contact_Period",
#             "SFT_SK": "SOFT_SKILLS",
#             "GEN": "GENDER",
#             "CTRY": "COUNTRY",
#             "JON_DT": "JOINING_DATE",
#             "SHFT_TM": "SHIFT_TIMINGS",
#             "PH": "PHONE",
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

#     import json

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

# return "\n".join(finalText)


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


#
# read multiple files with any functionality
# Import Module
# # import os

# # Folder Path
# path =r"C:/Users/admin/Downloads/New folder (7)/pdf"

# # Change the directory
# os.chdir(path)

# # Read text File


# def read_text_file(file_path):
#     with open(file_path, 'r') as f:
#         # print(f.read())
#         print(convert_pdf_to_txt(f.read()))


# # iterate through all file
# for file in os.listdir():
#     # Check whether file is in text format or not
#     if file.endswith(".pdf"):
#         file_path = f"{path}\{file}"

#         # call read text file function
#         read_text_file(file_path)
# #


# def pdf2img(pdf_file):
#     try:
#         images = pdf2image.convert_from_path(pdf_file)
#         for img in images:
#             img.save('./media/output.jpg', 'JPEG')
#             text = pytesseract.image_to_string(img)
#             return text


# import os
# import io
# from PIL import Image
# import pytesseract
# from wand.image import Image as wi
# import gc

# def Get_text_from_image(pdf_path):
#     pdf=pdf2image.convert_from_path(pdf_path)
#     # pdfImg=pdf.convert('jpeg')
#     imgBlobs=[]
#     extracted_text=[]
#     for img in pdfImg.sequence:
#         page=wi(image=img)
#         imgBlobs.append(page.make_blob('jpeg'))
#     for imgBlob in imgBlobs:
#         im=Image.open(io.BytesIO(imgBlob))
#         text=pytesseract.image_to_string(im,lang='eng')
#         extracted_text.append(text)
#     return ([i.replace("\n","") for i in extracted_text])


# 26-04-2021 for testing only not in use till date
# pdf to png conversion process
# def ppdf_pages(pdf_file):
#     pages = pdf2image.convert_from_path(pdf_file)
#     counter = 0
#     for page in pages:
#         file_name = 'page' + str(counter) + '.jpg'
#     # Save images to the same folder
#     page.save(file_name, 'JPEG')
#     # Open the file as an image
#     image_file = PIL.Image.open(file_name)
#     # Use tesseract to extract the text from the image
#     text = pytesseract.image_to_string(image_file)
#     # Print the contents to the console
#     print(text)
#     counter = counter + 1
#     return ''.join(text)
#     doc = nlp(str(text))
#     text_ = [entity.text for entity in doc.ents]
#     label = [entity.label_ for entity in doc.ents]
#     c = dict(zip(text_, label))
#             # print(c)
#             # print(pages,text)
#     return(c)
# end


#######
# json read
# # Folder Path
# testt = "C:/Users/admin/Documents/resume_parsing/resume"

# # Change the directory
# os.chdir(testt)

# # Read text File


# def read_text_file(file_path):
#     with open(file_path, 'r') as f:
#         print(f.read())

# # iterate through all file
# for file in os.listdir():
#     # Check whether file is in text format or not
#     if file.endswith(".txt"):
#         file_path = f"{testt}\{file}"

#         # call read text file function
#         read_text_file(file_path)
