import openpyxl 
import pandas as pd



wb = openpyxl.load_workbook('AuthorsListnContributions.xlsx')
# This workbook has only one sheet, therefore, we choose the active sheet directly
ws = wb.active

df = pd.DataFrame(ws.values)
d_df = df.to_dict()
first_name = d_df[0]
del first_name[0]
last_name = d_df[1]
del last_name[0]
email_address = d_df[2]
del email_address[0]
affiliation1 = d_df[3]
del affiliation1[0]
affiliation2 = d_df[4]
del affiliation2[0]
contribution = d_df[5]
del contribution[0]
authorship_order = d_df[6]
del authorship_order[0]
affiliation_list = {}
affiliation_order = 0
author_para = ""
author_with_affiliation = {}
for index in authorship_order:
    author_para += first_name[index] + " " + last_name[index]
    # print(item)
    # print(first_name[i] + "\t" + str(i) + "\t" + str(item[1]))
    # print(first_name[index] + " " + last_name[index])

    org1 = affiliation1[index]
    org2 = affiliation2[index]
    affiliation_superscript = ""
    if org1 is not None:
        if org1 not in author_with_affiliation:
            author_with_affiliation[org1] = [index]
        else:
            author_with_affiliation[org1].append(index)
        org1_index = list(author_with_affiliation.keys()).index(org1) + 1
        affiliation_superscript = str(org1_index)
    
    if org2 is not None:
        if org2 not in author_with_affiliation:
            author_with_affiliation[org2] = [index]
        else:
            author_with_affiliation[org2].append(index)
        org2_index = list(author_with_affiliation.keys()).index(org2) + 1
        if org1_index < org2_index:
            affiliation_superscript += ", " + str(org2_index)
        else:
            affiliation_superscript = str(org2_index) + ", " + affiliation_superscript
    author_para += "<sup>" + affiliation_superscript + "</sup>"#get_super(affiliation_superscript)
    author_para += ", "

i = 1
institutions = ""
for item in author_with_affiliation:
    # print(str(i) + "\t" + str(author_with_affiliation[item]) + "\t" + item)
    # print(get_super(str(i)) + item)
    institutions += "<sup>" + str(i) + "</sup>" + item + "<br/>" 
    i += 1

file = open("output.html", "w")

# Write HTML content
file.write("<html>")
file.write("<head>")
file.write("<title>Orders of the authors</title>")
file.write("</head>")
file.write("<body>")
file.write("<h1>Orders of the authors</h1>")
file.write("<p>" + author_para + "</p>")
file.write("<p>" + institutions + "</p>")
file.write("</body>")
file.write("</html>")

file.close()

print("The orders of the authors in your manuscript has been written in HTML file successfully.")