from os import listdir #list files
import xml.etree.ElementTree as ET #import xml to python
import re #regular expression
from scipy import misc #read image file
import os.path #check file existence
import re #regular expression
from pandas.io import sql #sql
import pymssql as mssql #sql
import Levenshtein #calculate distance, score of 2 sentences
from collections import OrderedDict # sort dictionary

const_default_neg_val = -1
const_default_page_number = 0
const_default_val = 1
const_default_str_null = "NULL"
const_default_reg_pattern = "\w+"
const_default_reg_pattern_number = "[\d]+[\d\.,]+" 
height_tolerance = 10
fileType = '1003'

class wdClass:
    l =0
    t=0
    r=0
    b=0
    line_t = 0
    count = 1
    text = ""

class ImageRange:
    def __init__(self, name):
        self.column = name
        self.center = []
        self.left = []
        self.top = []
        self.right = []
        self.buttom = []
        self.page = const_default_page_number
        self.reg_pattern = const_default_reg_pattern

class ImageBound:
    def __init__(self, name) :
        self.column = name
        self.l = const_default_val
        self.t = const_default_val
        self.r = const_default_val
        self.b = const_default_val

class ImageTable:
    def __init__(self, name, rows, row_map, page) :
        self.name = name
        self.rows = rows
        self.row_map = row_map
        self.page = page
        
class ImageCell:
    def __init__(self, name, reg):
        self.name = name
        self.reg_pattern = reg
        self.line_t = const_default_val

def getImagePosition(range, wordAddressMap):
    bound = ImageBound(range.column)
    for left in range.left:
        if left in wordAddressMap:
            bound.l = wordAddressMap[left].r
            break

    for right in range.right:
        if right in wordAddressMap:
            bound.r = wordAddressMap[right].l
            break

    for top in range.top:
        if top in wordAddressMap:
            bound.t = wordAddressMap[top].b
            break

    for buttom in range.buttom:
        if buttom in wordAddressMap:
            bound.b = wordAddressMap[buttom].t
            break

    return bound
    

#Computing the length of the LCS
def LCS(X, Y):
    m = len(X)
    n = len(Y)
    # An (m+1) times (n+1) matrix
    C = [[0 for j in range(n+1)] for i in range(m+1)]
    for i in range(1, m+1):
        for j in range(1, n+1):
            if X[i-1] == Y[j-1]: 
                C[i][j] = C[i-1][j-1] + 1
            else:
                C[i][j] = max(C[i][j-1], C[i-1][j])
    return C

#Reading out an LCS
def backTrack(C, X, Y, i, j):
    if i == 0 or j == 0:
        return ""
    elif X[i-1] == Y[j-1]:
        return backTrack(C, X, Y, i-1, j-1) +' '+ X[i-1]
    else:
        if C[i][j-1] > C[i-1][j]:
            return backTrack(C, X, Y, i, j-1)
        else:
            return backTrack(C, X, Y, i-1, j)

#page 1
Mortgage_Applied_for = ImageRange("Mortgage Applied for")
Mortgage_Applied_for.center.append("Applied");
Mortgage_Applied_for.center.append("VA");
Mortgage_Applied_for.center.append("FHA");
Mortgage_Applied_for.center.append("Conventional");
Mortgage_Applied_for.right.append("Agency");
Mortgage_Applied_for.right.append("Case");
Mortgage_Applied_for.top.append("joint");
Mortgage_Applied_for.top.append("credit");
Mortgage_Applied_for.top.append("Borrower");
Mortgage_Applied_for.buttom.append("Interest");
Mortgage_Applied_for.buttom.append("Rate")
Mortgage_Applied_for.buttom.append("Months")
Mortgage_Applied_for.buttom.append("Amortization")
Mortgage_Applied_for.buttom.append("Fixed")
Mortgage_Applied_for.page = 1

Amount =  ImageRange("Amount")
Amount.center.append("Amount")
Amount.top.append("Applied")
Amount.top.append("FHA")
Amount.top.append("VA")
Amount.top.append("Conventional")
Amount.top.append("Agency")
Amount.top.append("Case")
Amount.right.append("Interest")
Amount.right.append("Rate")
Amount.right.append("Months")
Amount.right.append("Amortization")
Amount.right.append("Fixed")
Amount.buttom.append("Subject")
Amount.buttom.append("Property")
Amount.buttom.append("appendress")
Amount.page = 1
Amount.reg_pattern = const_default_reg_pattern_number

Interest_Rate =  ImageRange("Interest_Rate")
Interest_Rate.center.append("Interest")
Interest_Rate.top.append("Applied")
Interest_Rate.top.append("FHA")
Interest_Rate.top.append("VA")
Interest_Rate.top.append("Conventional")
Interest_Rate.top.append("Agency")
Interest_Rate.top.append("Case")
Interest_Rate.left.append("Amount")
Interest_Rate.right.append("Months")
Interest_Rate.right.append("Amortization")
Interest_Rate.right.append("Fixed")
Interest_Rate.buttom.append("Subject")
Interest_Rate.buttom.append("Property")
Interest_Rate.buttom.append("appendress")
Interest_Rate.page = 1
Interest_Rate.reg_pattern =const_default_reg_pattern_number

Months =  ImageRange("Months")
Months.center.append("Months")
Months.top.append("Applied")
Months.top.append("FHA")
Months.top.append("VA")
Months.top.append("Conventional")
Months.top.append("Agency")
Months.top.append("Case")
Months.left.append("Interest")
#Months.left.append("Rate") # not unique
Months.left.append("Amount")
Months.right.append("Fixed")
Months.right.append("Amortization")
Months.buttom.append("Subject")
Months.buttom.append("Property")
Months.buttom.append("appendress")
Months.page = 1 
Months.reg_pattern = const_default_reg_pattern_number

Purpose =  ImageRange("Purpose")
Purpose.center.append("Purpose")
Purpose.center.append("Purchase")
Purpose.center.append("Construction")
Purpose.top.append("Legal")
Purpose.top.append("Description")
Purpose.top.append("attach")
Purpose.top.append("necessary")
Purpose.top.append("Built")
Purpose.right.append("Primary")
Purpose.right.append("Secondary")
Purpose.right.append("Investment")
Purpose.buttom.append("Liens")
Purpose.buttom.append("Present")
Purpose.buttom.append("Lot")
Purpose.buttom.append("Improvements")
Purpose.page = 1

Property_will_be = ImageRange("Property_will_be")
Property_will_be.center.append("Primary")
Property_will_be.center.append("Secondary")
Property_will_be.left.append("Construction")
Property_will_be.left.append("Purchase")
Property_will_be.left.append("Purpose")
Property_will_be.top.append("Built")
Property_will_be.top.append("Legal")
Property_will_be.top.append("Description")
Property_will_be.top.append("attach")
Property_will_be.top.append("necessary")
Property_will_be.buttom.append("Liens")
Property_will_be.buttom.append("Present")
Property_will_be.buttom.append("Lot")
Property_will_be.page = 1 

ImageRangeMap = {}
ImageRangeMap["Mortgage Applied for"]= Mortgage_Applied_for
ImageRangeMap["Amount"]= Amount
ImageRangeMap["Interest_Rate"]= Interest_Rate
ImageRangeMap["Months"]= Months
ImageRangeMap["Purpose"]= Purpose
ImageRangeMap["Property_will_be"]= Property_will_be

#page 3 
transaction_rows = (u'a. Purchase price',\
                   u'b. Alterations improvements repairs',\
                   u'c. Land acquired separately',\
                   u'd. Refinance incl debts to be paid off',\
                   u'e. Estimated prepaid items',\
                   u'f. Estimated closing costs',\
                   u'g. PMI MIP Funding Fee',\
                   u'h. Discount if Borrower will pay',\
                   u'i. Total costs add items through',\
                   u'j. Subordinate financing',\
                   u'k. Borrower closing costs paid by Seller',\
                   u'l. Other Credits explain',\
                   u'm. Loan amount',\
                   u'n. PMI MIP Funding Fee financed',\
                   u'o. Loan amount add m n',\
                   u'p. Cash from to Borrower')

transaction_rows_map = {}
for item in transaction_rows:
    row_map = ImageCell(item, const_default_reg_pattern_number)
    transaction_rows_map[item] = row_map

ImageTableMap = {}
ImageTableMap['transaction'] =ImageTable('transaction', transaction_rows, transaction_rows_map, 3)

fileDir = "\\\\54.164.134.136\\nfs\\Work\\tiwei\\1003\\"
XmlPath = "\\\\54.164.134.136\\nfs\\Work\\tiwei\\1003\\xml\\"
imagePath = "\\\\54.164.134.136\\nfs\\Work\\tiwei\\1003\\png\\"
ImageCut = "\\\\54.164.134.136\\nfs\\Work\\tiwei\\1003\\ImageCut\\"

#fileDir = "C:\\src\\files\\1003\\pdf"
#XmlPath = "C:\\src\\files\\1003\\xml"
#imagePath = "C:\\src\\files\\1003\\png"
#ImageCut = "C:\\src\\files\\1003\\result"

# determine which page we are going to process
process_ImageRange_page = []
for key, value in ImageRangeMap.iteritems():
    page = value.page
    if page not in process_ImageRange_page:
        process_ImageRange_page.append(page)
process_ImageRange_page = sorted(process_ImageRange_page)

process_ImageTable_page = []
for key, value in ImageTableMap.iteritems():
    page = value.page
    if page not in ImageTableMap:
        process_ImageTable_page.append(page)
process_ImageTable_page = sorted(process_ImageTable_page)


#list of pdf filenames
fileNames = []
for eachFile in listdir(fileDir):
    if eachFile.find("pdf") != -1:
        fileNames.append(eachFile.replace(".pdf", ""))
fileNames = sorted(fileNames)

#list of xml filenames
xmlFileNames = []
for eachFile in listdir(XmlPath):
    if eachFile.find("xml") != -1:
        xmlFileNames.append(eachFile)
xmlFileNames = sorted(xmlFileNames)

conn = mssql.connect(host='dev1db01', database='FileParsing', user='sa', password='SoFarSoGood.fk12')
conn.autocommit(True)
cursor = conn.cursor()

"""
for xmlFileName in xmlFileNames:
    print xmlFileName
    all_rows = sql.read_frame("select count(*) as cnt from  [FileParsing].[dbo].[xmlFiles] where FileName = '{0}'".format(xmlFileName), conn , coerce_float=True)
    
    #skip loading xml file into database if it exist in db
    if all_rows['cnt'][0] == 0:

        #parse xml file into db
        xmlFile = XmlPath +'\\'+ xmlFileName
        tree = ET.parse(xmlFile)
        root = tree.getroot()
        namespace = {'ab': 'http://www.scansoft.com/omnipage/xml/ssdoc-schema3.xsd'}
        paths = ["ab:page/ab:zones/ab:textZone/ab:ln","ab:page/ab:zones/ab:tableZone/ab:cellZone/ab:ln"]
        for path in paths:
            allLn = root.findall(path, namespace) 
            for Ln in allLn:
                word = wdClass() 
                word.line_t = int(Ln.get('t'))
                allWd= Ln.findall("ab:wd", namespace)
                for wd in allWd:
                    word.l = int(wd.get('l'))
                    word.t = int(wd.get('t'))
                    word.r = int(wd.get('r'))
                    word.b = int(wd.get('b'))
                    combine_text = '' if wd.text == None else wd.text
                    allRun = wd.findall("ab:run", namespace)
                    for run in allRun:
                        combine_text += run.text

                    try:
                        word.text = re.sub(r'[^a-zA-Z0-9\.,]','', combine_text)
                    except:
                        word.text = const_default_str_null
    
                    if word.text!= const_default_str_null:
                        text = "'{0}'".format(word.text)
                    else:
                        text = const_default_str_null
        
                    query = "insert into FileParsing.dbo.xmlFiles ([DocType],[FileName],[left],[top],[right],[buttom],[line_t],[text]) values('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', {7})".format(fileType, xmlFileName, word.l, word.t, word.r, word.b, word.line_t, text)
                    try:
                        cursor.execute(query)
                    except:
                        print query
"""

for fileName in fileNames: # for each file
    print fileName
    # for each column
    for key, value in ImageRangeMap.iteritems(): 
        xmlFileName = fileName + '_' + str(value.page).zfill(4) + '.xml'
        qry = "select replace(text, '.', '') as new_text , * from  [FileParsing].[dbo].[xmlFiles] where FileName = '{0}' and text is not null".format(xmlFileName)
        good_rows = sql.read_frame(qry, conn , coerce_float=True)
        if good_rows.values.shape[0] == 0:
            continue

        wordAddressMap = {}

        for row_number in range(0,good_rows.shape[0]):
            word = wdClass() 
            word.text = good_rows['new_text'][row_number]
            if word.text not in wordAddressMap:
                word.line_t = good_rows['line_t'][row_number]
                word.l = good_rows['left'][row_number]
                word.t = good_rows['top'][row_number]
                word.r = good_rows['right'][row_number]
                word.b = good_rows['buttom'][row_number]
                wordAddressMap[word.text] = word

        png_path = imagePath + '\\' + fileName + '.pdf-' + str(value.page).zfill(6) + '.png'
        if os.path.isfile(png_path) ==  True: 

            source = misc.imread(png_path) #read png 
            max_height = source.shape[0] # 0 is height
            max_width = source.shape[1] # 1 is width
            
            
            img = getImagePosition(value, wordAddressMap)
            if img.l > img.r and img.r != const_default_val :
                img.l =1 
            if img.t > img.b and img.b != const_default_val :
                img.t =1 
    
            width = max_width - img.l/10 if img.r == 1 else (img.r - img.l)/10
            height = max_height - img.t/10 if img.b==1 else  min(img.b/10 + height_tolerance, max_height) - img.t/10

            #cut image
            CroppedImage  = source[img.t/10: img.t/10+ height, img.l/10:img.l/10 + width]
            newImgPath = ImageCut + '\\' + fileName + '_' + img.column + ".png"
            misc.imsave(newImgPath,CroppedImage)

            query = "EXEC [FileParsing].[dbo].[display_xml_row] @doc_type = N'{5}', @fileName = N'{4}', @l = {0}, @t = {1}, @r = {2}, @b = {3}".format(img.l, img.t, img.l+width*10, img.t + height *10, xmlFileName, fileType)
            cursor.execute(query)

            #insert correponding data
            text_combine = ''
            for row in cursor:
                text_combine = text_combine + row[1] + ' '
            
            match = re.findall(value.reg_pattern, text_combine)        
            if(len(match) == 0): #no matching
                break 

            data = const_default_str_null 
            if value.reg_pattern == const_default_reg_pattern :
                data = " ".join(match)
                data = "'" + data + "'" 
            elif value.reg_pattern == const_default_reg_pattern_number : 
                data = match[0] 
                data = "'" + data + "'"

            query = "insert into FileParsing.dbo.xmlParsingResult ([FileName],[column_name],[text]) values('{0}', '{1}', {2})".format(xmlFileName, img.column, data)
            try:
                cursor.execute(query)
            except:
                print query


    for key, value in ImageTableMap.iteritems():
        all_rows = {}
        xmlFileName = fileName + '_' + str(value.page).zfill(4) + '.xml'
        qry = "EXEC [FileParsing].[dbo].[display_xml_row] @doc_type = N'{0}', @fileName = N'{1}'".format(fileType, xmlFileName)
        cursor.execute(qry)
        for row in cursor:
            all_rows[row[1]] = row[0] #key: text, value: line_t

        all_score = {}
        for data in value.rows:
            all_score[data] = -1    
            
        for key in all_rows:
            for data in all_score:
                ratio = Levenshtein.ratio(key, data)
                if ratio > all_score[data]:
                    all_score[data] = ratio
                    value.row_map[data].line_t = all_rows[key]
    
        # sort by line_t
        sorted_by_line_t = OrderedDict(sorted(value.row_map.items(), key=lambda x: x[1].line_t))
        sorted_index_array = []
        index_array = []

        # find the best matching rows in the table, longest common sequence
        row_number = 0
        for row in sorted_by_line_t:
            index = str(value.rows.index(row))
            sorted_index_array.append(index)
            index_array.append(str(row_number))
            row_number = row_number + 1 

        C = LCS(sorted_index_array, index_array)
        lcs_seq = backTrack(C, sorted_index_array, index_array, len(sorted_index_array), len(index_array))
        lcs_arr = lcs_seq.split()
        sorted_all_rows_by_line_t = OrderedDict(sorted(all_rows.items(), key=lambda x: x[1]))

        first_line = const_default_val
        last_line = const_default_val

        for index in range(len(lcs_arr)):
            cur_index_int = int(lcs_arr[index])
            cur_text = value.rows[cur_index_int]
            cur_line_t = value.row_map[cur_text].line_t
           
            nxt_line_t = const_default_str_null
            if index != len(lcs_arr) - 1 :
                nxt_index_int = int(lcs_arr[index+1])
                nxt_text = value.rows[nxt_index_int]
                nxt_line_t = value.row_map[nxt_text].line_t
            
            if index == 1: first_line = cur_line_t 
            if index == len(lcs_arr) - 1 : last_line = cur_line_t

            qry = "EXEC [FileParsing].[dbo].[display_xml_row] @doc_type = N'{0}', @fileName = N'{1}', @t = {2}, @b = {3}".format(fileType, xmlFileName, cur_line_t, nxt_line_t)
            cursor.execute(qry)
    
            #insert correponding data
            text_combine = ''
            for row in cursor:
                text_combine = text_combine + row[1]
            
            text_combine = text_combine.replace(" ", "")

            match = re.findall(const_default_reg_pattern_number , text_combine)        
            if(len(match) == 0): #no matching
                break 

            data = match[0] 
            data = "'" + data + "'"

            query = "insert into FileParsing.dbo.xmlParsingResult ([FileName],[column_name],[text]) values('{0}', '{1}', {2})".format(xmlFileName, cur_text, data)
            try:
                cursor.execute(query)
            except:
                print query

                """
        png_path = imagePath + '\\' + fileName + '.pdf-' + str(value.page).zfill(6) + '.png'
        if os.path.isfile(png_path) ==  True: 

            source = misc.imread(png_path) #read png 
            max_height = source.shape[0] # 0 is height
            max_width = source.shape[1] # 1 is width
            
            width = max_width
            height = (last_line - first_line )/10
            #cut image
            #CroppedImage  = source[first_line/10: first_line/10 + height, 1:max_width/2]
            #newImgPath = ImageCut + '\\' + fileName + '_' + value.name + ".png"
            #misc.imsave(newImgPath,CroppedImage)
            """