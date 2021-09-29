## Extract Data From PDF In Python

![MIT License](https://img.shields.io/github/license/shine-jayakumar/Covid19-Exploratory-Analysis-With-SQL)

  In this project, we are going to batch-convert pdf files to text and extract data without using PyPDF2/4. 

  We're going to achieve that by:
  - Using PDFtoText converter from XPdfReader to convert pdf files to text files
  - Using regular expressions to extract data
  - Performing data cleaning using pandas
  - Exporting to Excel file

## Why Not Use PyPDF2/4
  **Short Answer:** I got this error: 
  ```
  TypeError: object of type 'IndirectObject' has no len()
  ```
  
  **Long Answer:**  If PyPDF4 had worked I would never have had a chance to explore other ways. 
  I looked on [StackOverflow](https://stackoverflow.com/users/6711954/shine-j) however couldn't find a solution for this error. 
  Obviously, there had to be someone with the [same problem](https://stackoverflow.com/questions/66587056/typeerror-object-of-type-indirectobject-has-no-len) but there's no solution.
  
  I was not willing to manually copy and paste the information from 52 of my payslips.
  Isn't that what programs are used for? 
  
  
**Table of Contents**

- [Packages](#Packages "Packages")
- [Converting PDF To Text](#Converting-PDF-To-Text "Converting PDF To Text")
- [Script Link](#Script-Link "Script Link")


## Packages

- Pandas

  Check out the [requirements.txt](https://github.com/shine-jayakumar/Extract-Data-From-PDF-In-Python/blob/main/requirements.txt "requirements.txt")

## Converting PDF To Text
  Converting PDF to text using [Xpdf's pdftotext](http://www.xpdfreader.com/download.html "Xpdf's pdftotext") is really simple.

  Using this command-line tool we can batch-convert PDFs to text files.
  ```
  pdftotext source.pdf dest.txt
  ```
## Script Link
**Script Link:** [parse_payslips.py](https://github.com/shine-jayakumar/Extract-Data-From-PDF-In-Python/blob/main/parse_payslips.py "parse_payslips.py")
