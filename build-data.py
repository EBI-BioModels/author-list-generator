import requests

headers = {
    "Content-Type": "application/json"
}
params = {
    "format": "json",
    "resulttype": "core"
}
EUROPE_PMC = "https://www.ebi.ac.uk/europepmc/webservices/rest/search/query=ext_id:PMID%20src:med&resulttype=core&format=json"

PMID_SET = ["38863274", "38935577", "38863299", "38895901", "38950402", "38127458", "38932980", "38926517", "38924012", "36111857", "34878104", "33167871"]
for id in PMID_SET:
    new_url = EUROPE_PMC.replace("PMID", id)
    # print(new_url)
    res = requests.get(url=new_url, headers=headers)
    json_data = None
    if res.status_code == 200: 
        json_data = res.json()

    if json_data is not None:
        # print(json.dumps(json_data))
        author_list = json_data["resultList"]["result"][0]["authorList"]["author"]
        for author in author_list:
            first_name = author["firstName"]
            last_name = author["lastName"]
            if "authorAffiliationDetailsList" in author:
                affiliations = author['authorAffiliationDetailsList']['authorAffiliation']
                aff1 = affiliations[0]["affiliation"]
                aff2 = None
                if len(affiliations) >= 2:
                    aff2 = affiliations[1]["affiliation"]
                if aff2 is not None:
                    print(first_name + "\t" + last_name + "\t" + "test@email.com\t" + aff1 + "\t" + aff2)
                else:
                    print(first_name + "\t" + last_name + "\t" + "test@email.com\t" + aff1)
