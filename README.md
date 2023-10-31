# Invoice web scraping
## Python automation sample code to web scrape xlsx invoice files from a partner webapp

This project is just a demonstration of a similir project of mine, developed at the "Manchester Investimentos" financial department. Every part of this project is sample code which shows how to do the following:
* Search for debit notes in Outlook with pywin32.
* Read the debit notes using pdfplumber.
* Download invoices using selenium.
* Append invoices with pandas.
* Send file with pywin32.
* Delete files using os.

Every Tuesday the provider of the software sends us debit notes containing the invoice number which we use to get the extract of that invoice, as we control that file. As Manchester Investimentos have many branches in Brazil the invoices are segmented by CNPJ (like US EIN), we end up having many files resulting in a long process of donwloading and treating the data (as our partner doesn't have a optimized web app). So this automation does in 3 minutes what i would do in 30! Saving me time and keeping me away from boring work!

### See it working!
https://github.com/caio-pavesi/autodownload_invoices_PAYTRACK/assets/91769150/1ff9437e-6ddb-4e68-afe2-1aa71b5d7963
