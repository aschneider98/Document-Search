import os   # for accessing folder of documents
import csv  # for outputting matches to csv
import sys  # for showing progress in run output
from datetime import datetime   # for finding runtime of program
import docx2txt     # for docx files
import fitz     # for pdf files
import pandas as pd     # for xlsx files (could probably be accomplished with openpyxl)
import re


# mute errors from poorly formed pdfs
fitz.TOOLS.mupdf_display_errors(False)

# setting path for folder with documents to search through and path for csv to write to
docPath = 'C:\\Users\\schne\\PycharmProjects\\DocSearcher\\Docs\\pdfs-master'
csvPath = 'C:\\Users\\schne\\PycharmProjects\\DocSearcher\\matches.csv'

# capturing all filenames into list
files = os.listdir(docPath)

# creating terms to search for in documents
# searchTerms = {'the'}
searchTerms = {'parallel', 'the', 'computer', 'and', 'Fairbanks', 'Imputationvar', 'definitelynothere', 'C++', 'of',
               'this', 'is', 'a', 'test', 'more', 'words', 'please', 'seventeen', '18', 'many', 'end'}

# creating blank csv for answers to go to
with open(csvPath, 'w') as f:
    f.close()


# function for searching that loops through all documents in folder
def search_docs(documents, searchterms):
    for doc in documents:
        terms = searchterms.copy()
        # print progress percent based on index of current file compared to number of all files
        percent = (documents.index(doc) / len(documents)) * 100
        sys.stdout.write('\rProgress: %d%%' % percent)
        sys.stdout.flush()
        # logic for handling each type of file
        if doc.endswith('.docx'):
            search_docx(doc, terms)
        elif doc.endswith('.pdf'):
            search_pdf(doc, terms)
        elif doc.endswith('.xlsx'):
            search_xlsx(doc, terms)


# function for searching docx files
def search_docx(document, terms):
    # convert docx to string with all text from file
    text = docx2txt.process(docPath + '\\' + document).split()
    # search for each term in the file and write matches to csv
    ans = [document]
    for word in text:
        if len(terms) == 0:
            break
        elif word in terms:
            ans.append(word)
            terms.remove(word)
    if len(ans) > 1:
        write_csv(csvPath, ans)


# function for searching pdf files
def search_pdf(document, terms):
    # use fitz (a.k.a. mupdf) to process pdf
    pages = []
    ans = [document]
    with fitz.open(docPath + '\\' + document) as myPDF:
        for page in myPDF:
            pages.append(page.getText())
        text = ''.join(pages).split()
        for word in text:
            if len(terms) == 0:
                break
            elif word in terms:
                ans.append(word)
                terms.remove(word)
        if len(ans) > 1:
            write_csv(csvPath, ans)


# function for searching xlsx files
def search_xlsx(document, terms):
    # convert excel sheet to pandas dataframe
    pages = []
    ans = [document]
    xlsx = pd.ExcelFile(docPath + '\\' + document)
    sheets = pd.read_excel(xlsx, sheet_name=None)
    for sheet in sheets:
        pages.append(sheets[sheet].to_string())     #Doesnt work for headers
    text = ''.join(pages).split()
    for word in text:
        if len(terms) == 0:
            break
        elif word in terms:
            ans.append(word)
            terms.remove(word)
    if len(ans) > 1:
        write_csv(csvPath, ans)


# function for writing matches to csv
def write_csv(path, text):
    with open(path, 'a') as f:
        writer = csv.writer(f)
        writer.writerow(text)


# run program and record / print run time
start = datetime.now()
search_docs(files, searchTerms)
totalTime = datetime.now() - start
print()
print('Duration: {}'.format(totalTime))
