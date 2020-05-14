from PyPDF3 import PdfFileReader
import re
import os
import pandas as pd
import xlrd
import csv

import docx
from pathlib import Path

#list of fields to match
pii = ['Family Name', 'Given Name', 'Date of Birth', 'Tax File Number', 'Phone Number','Mobile Number', 'Email Address']

#
# Open Directories and Process files:
# Description: This function scan for the file types in
#              the directory and call the appropriate function
#              to scan the file
#
def opendirforprocessingfiles():
    #base path of where the files are on my computer welcome to change based your
    #environment
    baseDirectory = Path(
        'C:/Users/archi/Dropbox/UniversityofAuckland/2012/SOFTENG206/Project/asua006_project_CW/Source/insane/challenge')
    filesBasePath = baseDirectory.iterdir()

    for fileName in os.listdir(baseDirectory):
        if (fileName.endswith('.pdf')):#processing pdf files read from the path
            fileSplit = fileName.split('.')
            fileNamePDF = fileName
            filePDFExtension = fileSplit[1]
            openFiles(filePDFExtension, fileNamePDF)
        if(fileName.endswith('.xlsx')): #processing xlsx files read from the path
            fileSplit = fileName.split('.')
            fileNameXLS = fileName
            fileXLSXExtension = fileSplit[1]
            openFiles(fileXLSXExtension, fileNameXLS)
        if (fileName.endswith('.docx')):#processing docx files read from the path
            fileNameDOCX = fileName
            fileSplit = fileName.split('.')
            fileDocxExtension = fileSplit[1]
            openFiles(fileDocxExtension, fileNameDOCX)
        if(fileName.endswith('.csv')):#processing csv files read from the path
            fileNameCSV = fileName
            fileSplit = fileName.split('.')
            fileCSVExtension = fileSplit[1]
            openFiles(fileCSVExtension, fileNameCSV)


#
# Managed Files
# Description: The function distribute the processing of files to different functions
# for processing different fileformat such as PDF. DOCX, XLSX and CSV
#
def openFiles(fileformat ="", fileName =""):

    if(fileformat == "pdf"):
        pdfFileHandler = PdfFileReader(open(fileName, 'rb'))
        processPDFFiles(pdfFileHandler)
    if(fileformat == "xlsx"):
        xlsxdoc = xlrd.open_workbook(fileName)
        processXlsxFiles(xlsxdoc)
    if(fileformat == "docx"):
        docxFileHandler = docx.Document(fileName)
        processDocxFiles(docxFileHandler)
    if(fileformat == "csv"):
        csv.register_dialect('myDialect', delimiter=',', skipinitialspace=True)
        with open(fileName, 'r') as csvfile:
            csvreader = csv.DictReader(csvfile)
            processCSVFiles(csvreader)


def processCSVFiles(reader):
    datasetcsv = []# stroing data fetched from the csv file
    dataset = []
    for field in reader:
        datasetcsv.append(dict(field))

    piiindex = 0
    lenofpii = len(pii) - 1

    lastname = ""
    fullname = ""

    for index in datasetcsv:#traversing the list to get dictionary data
        dicpiival = {}
        for key, value in index.items(): #accessing the dictionary key-value pairs
            if key == pii[piiindex] and isValueEmptyCSV(value):
                # print(key, " :", value)
                if(key == 'Family Name'):
                    lastname = value
                if key == 'Given Name':
                    fullname = value + " " +lastname
                    dataset.append(fullname)
                    # print(fullname)
                if(key == 'Date of Birth'):
                    dicpiival.update({key: value})
                if (key == 'Tax File Number'):
                    dicpiival.update({key: value})
                if (key == 'Phone Number'):
                    dicpiival.update({key: value})
                if (key == 'Mobile Number'):
                    dicpiival.update({key: value})
                if (key == 'Email Address'):
                    dicpiival.update({key: value})
                if(piiindex >= lenofpii): # discontinue when the end of pii is read
                    continue
                piiindex += 1

        #reset index of pii to 0
        if piiindex == lenofpii:
            piiindex = 0
        dataset.append(dicpiival)

#
# Processing DOCX files for PII information
#
def processDocxFiles(docx):
    #print(filehandler)
    tfncount = 0
    namecount = 0
    rawdata = []
    taxfilenum = ""


    data = []
    keys = None

    count = 0
    for table in docx.tables:
        for i, row in enumerate(table.rows):
        # text = (cell.text for cell in row.cells)
            for cell in row.cells:
                if cell.text != "":
                    rawdata.append(cell.text)

    str = "tax withheld box you must lodge a tax return. If no tax"
    index = 0
    # print(index)
    # print(len(rawdata))

    for i in range(0, len(rawdata)):
        if i == rawdata.index("tax withheld box you must lodge a tax return. If no tax"):
            index = i + 19
            print(index)
            print(rawdata[index])




#
# Processing PDF files for PII information
#
def processPDFFiles(filehandler):

    pages = filehandler.numPages - 1;
    tfnoccurence = 0
    addresscount = 0
    piidata = []


    for i in range(0,pages):
        pdfpg = filehandler.getPage(i)
        pdfPageContents = pdfpg.extractText()

        findSeq = re.findall('\S+', pdfPageContents)

        #search for TFN occurence
        searchtfn = re.search('\d+\W\d+', pdfPageContents)

        #check for variations of different addresses such PO Box and house number
        addresspo = re.search("PO\sB\w+\s\d+", pdfPageContents)
        addressst = re.search("\d+\W+\s[A-Z]\w+\s[A-Z]\w+", pdfPageContents)
        if searchtfn or addresspo or addressst:
            fullname = findSeq[61] + " " + findSeq[62]
            tfnoccurence = tfnoccurence + 1
            addresscount = addresscount + 1
            temp = {"Full Name": fullname,"TFN Count": tfnoccurence, "Address Count": addresscount}

            piidata.append(temp)

        tfnoccurence = tfnoccurence - 1#reset the count for TFN
        addresscount = addresscount - 1#reset the count for address


#
# Processing XLSX files for PII information
#
def processXlsxFiles(filehandler):
    docxlsx = filehandler

    #PII data set
    dataset = []

    #list of PII fields on the excel sheet
    lisofheadings = []

    #xcel sheet being read in
    sheet = docxlsx.sheet_by_index(0)
    numrows = sheet.nrows
    numcols = sheet.ncols

    for i in range(1, numrows):
        for j in range(1, numcols):
            colheadings = sheet.cell_value(0, j)
            data = sheet.cell_value(i, j)
            cellval = sheet.cell_type(i,j)#type of values in each cell

            # check against the list define above & data is in the cell
            if pii[0] == colheadings and data != "":
                familyname = data
            if pii[1] == colheadings and data != "":
                firstname = data
            if checkHeadingsAndValuesInExcelSheet(colheadings, data,i,j):
                if(len(lisofheadings) == 4):
                    continue
                lisofheadings.append(colheadings)

        fullname = firstname + " " + familyname
        dataset.append(fullname)
        dataset.append(lisofheadings)

#
# Check for certain PII fields header of the excel sheet
# being processed
#
def checkHeadingsAndValuesInExcelSheet(heading, data, row, col):
    if pii[2] == heading and data != 0:
        return True
    if pii[3] == heading and data != 0:
        return True

    if pii[4] == heading and data != "":
        return True

    if pii[5] == heading and data != "":
        return True

    if pii[6] == heading and data != "":
        return True

def listOfTuples(l1, l2):
    return list(map(lambda x, y:(x,y), l1, l2))

#
# Check that the cell value is not empty
#
def isValueEmptyCSV(value):
    if not value == 0 or value == "":
        return True
    else:
        return False

#
#
#
def csvListContains(checkdata):
    check = False
    for piidata in pii:
        if piidata == checkdata:
            check = True
            return check
    return check

def main():
    opendirforprocessingfiles()

if __name__ == "__main__":
    main()



