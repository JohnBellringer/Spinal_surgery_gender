#IMPORT PORTION
import openpyxl
from Bio import Entrez
import gender_guesser.detector as gender
from datetime import datetime
from dateutil.relativedelta import relativedelta

#SETTING UP

# Set up the Entrez API keys and email address (will be in a file-format in the future)
#print("Please, insert the name of the file from which we can take the API email and the API key")
Entrez.email = "luca.mascaro00@gmail.com"
Entrez.api_key = "85a0c7eeaa15bb56735b36487500a96b3208"
counter = 0
# Set up the gender detector (from Gender-guesser)
gd = gender.Detector()

# Parse the results and save them to an Excel file
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["DOI", "First Author Name", "First Author Female", "First Author Male"," First Author Nationality", "Last Author Name", "Last Author Female", "Last Author Male", "Last Author Nationality", "Year", "Journal"])

#INPUT PORTION (with related setting up and conversions)

#Set up the name for the Excel file

#print("Please, insert the name that you would like for the Excel file to save the research")
#nameExcel = input()
nameExcel = "String 7.xlsx"

# Set up the search query parameters (will be in an input-format in the future)
#print("Please, insert the query parameters of the research")
#search_query = input()
search_query = ('"Spinal Fusion" AND "Neurosurgery"')



#Set up the language (will be in an input-format in the future)
#print("Please, insert the language(s) of the publications for the research")
#language = input()
language = "eng"

# Definition of the start and the end date (will be in an input-format in the future)
#print("Please, insert the start date that you would like for the research")
#start_date = input()
#print("Please, insert the end date that you would like for the research")
#end_date = input()

start_date = "2017/09/01"
end_date = "2022/08/31"
#start_date = "2017/09/01"
#end_date = "2022/08/31"

# Convert the start and end dates to the required format
start_date = datetime.strptime(start_date, "%Y/%m/%d").strftime("%Y/%m/%d")
end_date = datetime.strptime(end_date, "%Y/%m/%d").strftime("%Y/%m/%d")

# Set up of the date for the following while loop
start_date_temp = start_date
end_date_temp = (datetime.strptime(start_date_temp, '%Y/%m/%d') + relativedelta(days=1)).strftime('%Y/%m/%d')

#WHILE LOOP

# start the cycle for the acquisition of the data, occurring every month
while end_date_temp <= end_date:

    # ENTREZ PORTION

    # Use Entrez to search for the relevant publications
    fetch_handle = Entrez.esearch(db="pubmed", term=search_query, mindate=start_date_temp, maxdate=end_date_temp,lang=language)
    fetch_results = Entrez.read(fetch_handle)
    fetch_handle.close()

    # Use Entrez to retrieve the full records for the relevant publications only if the ID list is not empty
    #if id_list:
    #    fetch_handle = Entrez.efetch(db="pubmed", id=id_list, rettype="xml", retmode="text")
    #    fetch_results = Entrez.read(fetch_handle)
    #    fetch_handle.close()

    # Use Entrez to retrieve the full records for the relevant publications
    id_list = fetch_results["IdList"]
    if id_list:
        fetch_handle = Entrez.efetch(db="pubmed", id=id_list, rettype="xml", retmode="text")
        fetch_results = Entrez.read(fetch_handle, validate=False)
        fetch_handle.close()
        # Parse the data, skipping tag validation
        #parsed_data = Entrez.read(fetch_results, validate=False)
        Entrez.parse(fetch_results, validate=False)

        #fetch_handle = Entrez.efetch(db="pubmed", id=id_list, rettype="xml", retmode="text", retmax=count)
        #fetch_handle = Entrez.efetch(db="pubmed", id=id_list, rettype="xml", retmode="text")
        #fetch_record = Entrez.read(fetch_handle)
        #count = int(fetch_record['Count'])

        #Define a For cycle for the writing in the Excel file of the single publications (with first/last authors, their gender, ...)
        for article in fetch_results["PubmedArticle"]:
            try:
                # Get the first and last author's name and nationality
                author_list = article["MedlineCitation"]["Article"]["AuthorList"]
                if isinstance(author_list, list):
                    author = author_list[0]
                    lastAuthor = author_list[-1]
                else:
                    author = author_list
                    lastAuthor = author_list


                # Obtain the name the first author
                first_name = author.get("ForeName", "")
                last_name = author.get("LastName", "")
                #first_auth_nationality = author.get("AffiliationInfo", [])[-1].get("Country", "") if "AffiliationInfo" in author and author["AffiliationInfo"] else ""

                #Nationality of the affiliation of the first author
                first_auth_nationality = ''
                first_auth_affiliations = author.get("AffiliationInfo", [])
                for affiliation in first_auth_affiliations:
                    first_auth_affiliation_string = affiliation.get("Affiliation", "").strip()
                    if "Electronic address:" in first_auth_affiliation_string:
                        first_auth_affiliation_string = first_auth_affiliation_string.replace("Electronic address:", "")
                    country = first_auth_affiliation_string.split(",")[-1].strip().split(" ")[0].rstrip(";").rstrip(".")
                    if country:
                        affiliation_parts = first_auth_affiliation_string.split(",")
                        if len(affiliation_parts) > 1:
                            next_words = affiliation_parts[-1].strip().split(" ")
                            if len(next_words) > 1:
                                next_word = next_words[1]
                                country += " " + next_word
                            first_auth_nationality = country
                            break

                # Obtain the info of the last author
                first_name_last = lastAuthor.get("ForeName", "")
                last_name_last = lastAuthor.get("LastName", "")
                last_auth_nationality = lastAuthor.get("AffiliationInfo", [])[-1].get("Country", "") if "AffiliationInfo" in lastAuthor and lastAuthor["AffiliationInfo"] else ""

                # Nationality of the affiliation of the last author
                last_auth_nationality = ''
                last_auth_affiliations = lastAuthor.get("AffiliationInfo", [])
                for affiliation in last_auth_affiliations:
                    last_auth_affiliation_string = affiliation.get("Affiliation", "").strip()
                    if "Electronic address:" in last_auth_affiliation_string:
                        last_auth_affiliation_string = last_auth_affiliation_string.replace("Electronic address:", "")
                    country = last_auth_affiliation_string.split(",")[-1].strip().split(" ")[0].rstrip(";").rstrip(".")
                    if country:
                        affiliation_parts = last_auth_affiliation_string.split(",")
                        if len(affiliation_parts) > 1:
                            next_words = affiliation_parts[-1].strip().split(" ")
                            if len(next_words) > 1:
                                next_word = next_words[1]
                                country += " " + next_word
                            last_auth_nationality = country
                            break

                # Get the DOI, year, and journal
                doi_list = article["PubmedData"]["ArticleIdList"] if "ArticleIdList" in article["PubmedData"] else []
                doi = next((i for i in doi_list if i.attributes.get("IdType") == "doi"), article["MedlineCitation"]["PMID"])
                year = article["MedlineCitation"]["Article"]["Journal"]["JournalIssue"]["PubDate"]["Year"] if "Year" in article["MedlineCitation"]["Article"]["Journal"]["JournalIssue"]["PubDate"] else ""
                journal = article["MedlineCitation"]["Article"]["Journal"]["Title"] if "Title" in article["MedlineCitation"]["Article"]["Journal"] else ""

                # Guess the gender of the first author
                gender_guess_first = gd.get_gender(first_name)
                if gender_guess_first == "male":
                    male_first = "1"
                    female_first = "0"
                elif gender_guess_first == "female":
                    male_first = "0"
                    female_first = "1"
                else:
                    male_first = "X"
                    female_first = "X"

                # Guess the gender of the last author
                gender_guess_last = gd.get_gender(first_name_last)
                if gender_guess_last == "male":
                    male_last = "1"
                    female_last = "0"
                elif gender_guess_last == "female":
                    male_last = "0"
                    female_last = "1"
                else:
                    male_last = "X"
                    female_last = "X"

                # Add the results to the Excel file
                ws.append([doi, f"{first_name} {last_name}", female_first, male_first, first_auth_nationality, f"{first_name_last} {last_name_last}", female_last, male_last, last_auth_nationality, year, journal])
                # Save the Excel file (the name of the Excel file will be in an input-format in the future)
                wb.save(nameExcel)

            except KeyError:
                #counter=counter+1
                continue

    else:
        # Handle the case when the ID list is empty
        # For example, you can print an error message or skip further processing.
        print("No publications found for the following day.")

    # Update the start and end dates (for the while cycle), adding one month
    start_date_temp = (datetime.strptime(start_date_temp, '%Y/%m/%d') + relativedelta(days=1)).strftime('%Y/%m/%d')
    end_date_temp = (datetime.strptime(end_date_temp, '%Y/%m/%d') + relativedelta(days=1)).strftime('%Y/%m/%d')

    print(" ")
    print("Date:", start_date_temp)
    print("Search Query:", search_query)
    print("Time filter:", start_date, end_date)
    print("ID List:", id_list)

    # Save the Excel file (the name of the Excel file will be in an input-format in the future)
    wb.save(nameExcel)
    #print(counter)


    
