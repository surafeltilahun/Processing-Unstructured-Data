# Processing unstructured data

This algorithm is inteded to process and convert unstructured data (PDF), with different format into semi-structured data (CSV). The algorithm is able to read-in PDF files, remove headers and footers, and merge rows then save the files in CSV format. The algorithm also checks if the data has been succesfully extracted from PDF files and if so, the source files from the Source folder will be moved to Archive folder, which enables us to keep tracking of the files that are succesfully processed or not by the algorithm. 

# Requirment

The algorithm requires column names and footers as a parameter and API from PDFTables (https://pdftables.com/) to convert the PDFs to Excel. It also requires four folders; Source, Converted, Output and Archive.


Source - is a folder that should contain the original PDF files

Converted - is a folder that contains the converted Excel files

Output - should comprise files that have been cleaned and saved in CSV format

Archive - this folder holds the original PDF files that have been successully processed by the algorithm (moved from source folder)




