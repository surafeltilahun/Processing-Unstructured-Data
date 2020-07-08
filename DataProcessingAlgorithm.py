#Import libraries
import pandas as pd
import numpy as np
import glob as g
import fnmatch as fm
import os as os
import pdftables_api as api
import shutil
import dateutil.parser as dparser
import re


#set the path to the files
sourcePath = r"*********************************************.*"
gripSourcePath = sourcePath[0:(sourcePath.rfind('\\'))]
destinationPath = r"************************************"
converterPath = r"********************************************"
outputPath =  r"*********************************************"

#collect all files in source folder
def collect_datafiles(a,b):
    try:
        b = g.glob(a)
    except:
        print("Error - While collecting InputFiles")
    return b

#convert the pdfs into excel using API from PDFTables
def convert_PDFtoExcel(a,b):
    try:
        conversion = api.Client('**********')
        conversion.xlsx_multiple(a,b)
    except:
        print("Error - While converting Source file -"+a+" from pdf to excel")
    return b

#move the files to Archive folder
def MoveToArchieve(a,b,c):
    s = a+"\\"+c
    d = b+"\\"+c
    if s:
        if os.access(b, os.W_OK):
            shutil.move(s,d)
        else:
            print("The access for moving file from source to destination was denied - ")
    else:
        print("System Admin - Please check the file path of source and archive")   
        
#import excel files from Converted folder
def load_inputfile(a,b):
    try:
        b = pd.read_excel(a,sheet_name=None)
    except:
        print("Error - While Processing Inputfile")
    return b

#remove an empty columns
def removeEmptyColumns(a,b):
    try:
        findEmptyColumns = [col for col in a.columns if a[col].isnull().all()]
        if (len(findEmptyColumns) > 0): 
            a.drop(findEmptyColumns, axis=1, inplace = True)
        b = a
        return b
    except:
        print("Error - While removing Blank columns")

#remove the NAs and replace them with an empty cell
def replaceNanValues(a,b):
    try:
        b = a.replace(np.nan, '', regex=True)
    except:
        print("Error - While cleaning blank values")
    return b

#remove headers and footers
def removeHeadersandFooters(a,b):
    try:
        a = pd.DataFrame(a)
        i = 0
        findfooterindex = a[a.apply(lambda row: row.astype(str).str.contains('End|Report|report').any(), axis=1)].index
        b = a.astype(str)
        header_ = re.compile("^ID.*Start.*$|^Task.*Job.*$|^Trade.*Company.*$|^Start.*Finish.*$|^Critical.*Job.*$|^Date.*Activity.*$")
        findheaderindex = b[b.apply(lambda x : True if header_.search("".join(x)) else False, axis=1)].index
        print(findheaderindex)
        findheaderindex = list(findheaderindex)
        findfooterindex = list(findfooterindex)
        if findheaderindex:
            if not (0 in findheaderindex):
                findheaderindex.extend([i])
        findheaderindex.sort()
        findfooterindex.sort()
        if findheaderindex:
            a.drop(a.index[min(findheaderindex):(max(findheaderindex))], axis = 0, inplace=True)
        if findfooterindex:
            if (len(findfooterindex) > 1):
                a.drop(a.index[min(findfooterindex):max(findfooterindex)], axis = 0, inplace=True)
            elif (len(findfooterindex) == 1):
                a.drop(findfooterindex, axis = 0, inplace=True)
        b = a    
    except:
        print("Error - while removing header and footer")
    return b

#merge the rows where there is space in between two consecutive rows
def rowMerger(a,b):
    try:
       rule1 = lambda x: x not in ['']
       u = a.loc[a['column0'].apply(rule1) & a['column1'].apply(rule1)].index
       findMergerindexs = list(u)
       findMergerindexs.sort()
       a = pd.DataFrame(a)
       tabcolumns = pd.DataFrame(a.columns)
       totalcolumns = len(tabcolumns)
       b = pd.DataFrame(columns = list(tabcolumns))
       if (len(findMergerindexs) > 0):
           for m in range(len(findMergerindexs)):
               if not (m == (len(findMergerindexs)-1)): 
                   startLoop = findMergerindexs[m]
                   endLoop = findMergerindexs[m+1]
               else:
                   startLoop = findMergerindexs[m]
                   endLoop = len(a)
               listValues = []
               for i in range(totalcolumns):
                   value = ''
                   for n in range(startLoop,endLoop):
                       value = value + ' ' + str(a.iloc[n,i])
                   listValues.insert(i,(value.strip()))
               b = b.append(pd.Series(listValues),ignore_index = True)
       else:
            print("File is not having a row for merging instances - Please check the file manually for instance - ")
       return b
    except: 
        print("Error - While merging the rows")
    return b

#move the final cleaned version of the data to Azure Blob Storage
def movetoAzureBlob(a,b):
    accountName = "********"
    accountKey = "************************************************************"
    containerName = "**********"
    blobService = BlockBlobService(account_name=accountName, account_key=accountKey)
    blobService.create_blob_from_text(containerName, b, a)

#the main function
try:
    sourceFiles = collect_datafiles(sourcePath,0)
    if sourceFiles:
        for sfile in sourceFiles:
            grabtheFileFormat = sfile[(sfile.rfind('.')+1):]
            grabtheFile = sfile[(sfile.rfind('\\')+1):]
            if (grabtheFileFormat.lower() == 'pdf'):
                """
                convert_PDFtoExcel(sfile,0)
                """
                convertedFile = grabtheFile.replace(grabtheFileFormat,'xlsx')
                convertedFile = converterPath + "\\"+convertedFile 
            elif (grabtheFileFormat.lower() == 'xlsx'):
                convertedFile = sfile
            else:
                convertedFile = ''
                print("The processing file is not in pdf or excel format - On hold for furthur Development")
            if convertedFile:
                datafiles = collect_datafiles(convertedFile,0)
                if datafiles:
                    for file in datafiles:
                        actualData = load_inputfile(file,0)
                        print("Collect Required Data from the file  - "+grabtheFile)
                        maxPossibleColumns = 0
                        for tabs in actualData.keys():
                            tabdata = pd.DataFrame(removeEmptyColumns(actualData[tabs],0))
                            tabcolumns = pd.DataFrame(tabdata.columns)
                            totalcolumns = len(tabcolumns)
                            if maxPossibleColumns < totalcolumns: 
                                maxPossibleColumns = totalcolumns
                            else: 
                                maxPossibleColumns = maxPossibleColumns
                        column_names = ['column'+ str(i) for i in range(maxPossibleColumns)]
                        data = pd.DataFrame(columns = column_names) ## Sheet Specific
                        dataCleaner = pd.DataFrame(columns = column_names) ## Whole Excel File
                        for tabs in actualData.keys():
                            print("Process Excel sheet - Started - "+tabs)
                            requiredData = pd.DataFrame(removeEmptyColumns(actualData[tabs],0))
                            tabcolumns = pd.DataFrame(requiredData.columns)
                            totalcolumns = len(tabcolumns)
                            for i in range(totalcolumns):
                                data.iloc[:,i] = requiredData.iloc[:,i]
                            data.reset_index(drop=True, inplace=True)
                            data = replaceNanValues(data,0)
                            data.reset_index(drop=True, inplace=True)
                            data = removeHeadersandFooters(data,0)
                            data.reset_index(drop=True, inplace=True)
                            for index, row in data.iterrows():
                                dataCleaner = dataCleaner.append(row)
                            print("Process Excel sheet - Ended   - "+tabs)
                        dataCleaner.replace({r'\n': ''}, regex=True, inplace=True)
                        dataCleaner.reset_index(drop=True, inplace=True)
                        dataCleaner = rowMerger(dataCleaner,0)
                        dataCleaner.columns=dataCleaner.iloc[0,:]
                        dataCleaner=dataCleaner.iloc[1:]
                        dataCleaner=dataCleaner.reset_index(drop=True)
                        print(dataCleaner)
                        writerpath = outputPath + "\\"+ grabtheFile.replace(grabtheFileFormat,'csv')
                        dataCleaner.to_csv(writerpath, index = False)
                        MoveToArchieve(gripSourcePath, destinationPath, grabtheFile)
                        
                        print("Collected Required Data from the file - "+grabtheFile)
                else:
                    print("The data is not collected from the processing file. Please check the file- "+ grabtheFile)  
            else:
                print("The processing file is not converted - On hold for furthur Development - "+grabtheFile)                    
    else:
        print("There is no file for processing in the repository. Please check - "+sourcePath)
except:
    print("Error - While processing steps")
finally:
    print("Process Completed")
    
