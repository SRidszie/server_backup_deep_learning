test_dict = {"gfg": 10, "is": 15, "best": 20, "for": 10, "geeks": 20}

blank_dict = {
    "first_name": [["Sandeep", "Sandeep", "Yash", "Yash", "Mani", "Tuffy"]],
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

# unique = set()
# for d in test_dict:
#     t = tuple(d.iteritems())
#     unique.add(t)

# print([dict(x) for x in unique])


list = []  # create empty list
for val in blank_dict.values():
    if val in list:
        continue
    else:
        list.append(val)


s = set()
for dic in blank_dict:
    for val in dic.values():
        s.add(val)

print(s)

# temp = {val: key for key, val in blank_dict.items()}
# res = {val: key for key, val in temp.items()}

# result = {}

# for key, value in blank_dict.items():
#     if value not in result.values():
#         result[key] = value

# print(result)


# print(temp)
# Remove duplicate values in dictionary
# Using loop
temp = []
res = dict()
for key, val in blank_dict.items():
    if val not in temp:
        temp.append(val)
        res[key] = val


# res_list = []
# for i in range(len(blank_dict)):
#     if blank_dict[i] not in blank_dict[i + 1 :]:
#         res_list.append(blank_dict[i])

# res_list = [i for n, i in enumerate(blank_dict) if i not in blank_dict[n + 1 :]]
res_list = {frozenset(item.items()): item for item in blank_dict}.values()

# printing result
# print("The dictionary after values removal : " + str(res_list))
