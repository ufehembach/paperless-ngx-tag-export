# paperless-ngx-tagexporter

this script exports documents and json files based on TAGs in paperless-ngx.

it needs a folder export, to be set in the ini file.
Add your paparless details also to the ini file: URL and TOKEN

for each token with a corrosponding diretoryname in the export folder, all dokuments in pdf and in json will be exported, filename is the title of the entry in paperless.
it also creates an excel file with most important informations about the pdf/json, including tags and all custom fields. 
for custom fields of type currency it creates a sum formular above the header of the table in the excel sheet.
also the table will be formatedd a little bit SAP like (alternating rows have a light backround) 
the id in the first column is also a link to the document in paperless web gui.

good luck with it.
this was a quick hack with ChatGPT, needed for my tax exports
