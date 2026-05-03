import re
# sample_dict = {'job_exp': ('5yrs', '6 years', 'Experience range 7-9 Years')}
sample_dict = {'job_exp': [('certified in Data Science',)]}


sample_dict["job_exp"] = "".join(str(e) for e in sample_dict["job_exp"])
# print(sample_dict["job_exp"])




z = re.findall(r"\d+",str(sample_dict["job_exp"]))
print(z)
z.append("9999999999999999")
print(z)


exp_min = int()
for i in min(map(str, z)):
    exp_min = i
print(exp_min)    
# sample_dict["exp_min"] = exp_min






# exp_min = int()
# for i in list(map(int, re.findall(r"\d+", sample_dict["job_exp"]))):
#     exp_min = i
# print(min(exp_min ,i))
# exp_min = min(map(int, re.findall(r"\d+", ''.join(sample_dict["job_exp"]))))
# exp_min = min(map(int, re.findall(r"\d+", "".join(str(e) for e in sample_dict["job_exp"]))))
# print(re.findall(r"\d+", "".join(str(e) for e in sample_dict["job_exp"])))
# print(exp_min)
# a = ["abc"]
# a = [[0] if not arr else arr for arr in a]
# print(a)
# for n, i in enumerate(a):
#    if i == []:
#        a[n] = [0]
# print(a)       


# for i in range(0, len(a)):
#     if a[i]==[]:
#         a[i]=[0]
# # a = [ [1], [2], [] , [4] ]
# # [ l[0] if l else 0 for l in a ]  
# print(a) 
# print(max([[0] if not arr else arr for arr in a],10))


# a = []
# a.append(0)
# # for i in a:
#     a.append(10)
    # if a[i] == []:
    #     a[i] = 0
# print(a)
# a = sum(a,[])
# print(a)

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
