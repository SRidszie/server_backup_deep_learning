sample_dict = {"name": ["Sandeep", "Sandeep", "Riddhi", "Riddhi"],
               "age": [30,30],
               "birthdate": ["November", "November"]
               }


# print(type(sample_dict))
# print(sample_dict)
# print(sample_dict.keys())
# print(type(sample_dict["name"]))
print(list(set(sample_dict["name"])))
# print(sample_dict["name"])
# print(sample_dict.items())







import re
sample_dict = {'job_exp': ('5yrs', '6 years', 'Experience range 7-9 Years')}

sample_dict["job_exp"] = "".join(str(e) for e in sample_dict["job_exp"])

# exp_min = int()
# for i in list(map(int, re.findall(r"\d+", sample_dict["job_exp"]))):
#     exp_min = i
# print(min(exp_min ,i))
exp_min = min(map(int, re.findall(r"\d+", ''.join(sample_dict["job_exp"]))))
print(exp_min)


# sample_dict["exp_min"] = exp_min

# print(sample_dict)

# print(sample_dict)
# exp_min = min(sample_dict.values())


# def high_and_low(numbers):
#     #    split numbers based on space     v
#     numbers = [int(x) for x in numbers.split()]
#     #           ^ type-cast each element to `int`
#     return max(numbers), min(numbers)

# print(high_and_low("3yrs, 5 years, Experience range 7-8 Years"))
