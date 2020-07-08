# Processing unstructured data

This project is inteded to process and convert unstructured data (PDF) into semi-structured data (CSV). The unstructured data have different sources and there is not similarity among them. The algorithm is able to read in PDF files, remove headers and footers, and merge rows then save the files in CSV format. The algorithm also checks if the data has been succesfully extracted from PDF files and if so, the source files from source folder will be moved to Archive folder, which enables us to keep track of the files that are succesfully processed or not. 

# Requirment

The algorithm requires column names and API from PDFTables (https://pdftables.com/) to convert the PDFs to Excel. 




