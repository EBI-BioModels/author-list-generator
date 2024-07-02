import openpyxl, random
import pandas as pd


FILE_NAME = "AuthorsListnContributions.xlsx"
wb = openpyxl.load_workbook(FILE_NAME)
# This workbook has only one sheet, therefore, we choose the active sheet directly
ws = wb.active

def generate_author_list():
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


def customise_data():
    # as of writing this comment, A2:A84 -> Authors First name, B2:B84 --> Last name
    list = generate_random_sequence()
    df = pd.DataFrame(ws.values)
    ddf = df.to_dict()
    first_name = ddf[0]
    del first_name[0]
    last_name = ddf[1]
    del last_name[0]
    j = 2    
    for i in list:
        print(i)
        print(first_name[i])
        print(last_name[i])
        ws["A" + str(j)] = first_name[i]
        ws["B" + str(j)] = last_name[i]
        j += 1
    wb.save(filename=FILE_NAME)


def print_rows():
    for row in ws.iter_rows(values_only=True):
        print(row)


def generate_random_sequence():
    # as of writing this comment, the sheet has A2:A84 -> Authors First name, B2:B84 --> Last name
    my_list = range(1,84)
    random_sample = random.sample(my_list, 83)
    
    return random_sample


def generate_with_random_names():
    print_rows()
    generate_random_sequence()
    customise_data()
    generate_author_list()
    print_rows()


if __name__ == "__main__":
    # generate_with_random_names()
    generate_author_list()