# WZ_finder
A Script searches for modules in WZ documents (excel workbook) then puts report to stdout and saves result in csv file.

WZ is abbreviation of Wydanie Zaopatrzenia in polish (material release receipt). It contains a list of names, quantities and serial numbers of electronic modules. Excel workbook conatins set of WZ documents, each document on separate sheet.
Modules can be sold, loaned, returned to client after repair. If there is a need to search for specific module (sold, loaned, returned to client) we can use WZ_finder to generate report. The report contains name of module, list of WZ documents numbers, serial numbers and quantities in particular WZ document. It is helpful for verification of operations or making statistics.

Usage example:
>python WZ_finder.py "WZ_documents.xls" "Module1" "sale"
