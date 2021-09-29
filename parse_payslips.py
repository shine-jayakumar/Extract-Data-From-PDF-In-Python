import os
import subprocess
import sys
import re
import pandas as pd
from datetime import datetime
import logging

# Setting up logger

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter("%(asctime)s:%(name)s:%(lineno)d:%(levelname)s:%(message)s")
file_handler = logging.FileHandler(f'parse_payslip_info_log_{datetime.now().strftime("%d_%m_%Y__%H_%M_%S")}.log')
file_handler.setFormatter(formatter)

stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setFormatter(formatter)

logger.addHandler(file_handler)

if len(sys.argv) > 1:
    if "--verbose" in sys.argv or "-v" in sys.argv:
        logger.addHandler(stdout_handler)




def convert_pdf_to_text():
    '''
    convert_pdf_to_text()

    Converts all pdf files in the current directort to text file

    Output directory: .\converted_pdfs

    '''
    # directory to hold converted text files
    logger.info("Creating converted_pdfs directory")
    if os.system("mkdir converted_pdfs") != 0:
        logger.error("Failed to create converted_pdfs directory")
        sys.exit()
    
    # list of pdf files in dir/sub-dir and save them to a text file
    logger.info("Gathering list of full path to pdf files in the current directory/sub-directory")
    os.system("dir /s /b *.pdf > allpdf.txt")
    
    list_fnames = []
    
    # put \n seperated file path in a list
    logger.info("Saving pdf file names to a list")
    try:
        with open('allpdf.txt', 'r') as fh:
            list_fnames = list(fh.read().split('\n'))
    except FileNotFoundError:
        logger.error("Unable to open file: addpdf.txt")
        sys.exit()
        
    err_count = 0
    

    # converting files one by one
    logger.info("Generating text files from pdf")
    for fname in list_fnames:
        if fname:
            target_text_fname = f"{get_fname_without_ext(fname)}.txt"
            target_text_path = os.path.join('.\converted_pdfs',target_text_fname)
            ret = subprocess.run(['bin64\pdftotext.exe', fname, target_text_path], capture_output=True)
            if ret.returncode != 0:
                logger.error(f"Error converting: {target_text_path}")
                err_count += 1
    # saving list of converted text files
    logger.info("Gathering list of text file names")
    os.system("dir converted_pdfs\*.txt /b > alltexts.txt")
    
    return err_count

def get_fname_without_ext(fname):
    '''
    get_fname_without_ext(fname)

    Returns the filename (without extension) from a filepath

    Ex: Return 'ebook' from d:\pdffiles\ebook.pdf 

    '''
    match = ""
    #pattern = re.compile(r'(Payslip_.+)(.pdf)')
    match = re.search(r'(Payslip_.+)(.pdf)',fname)
    if match:
        return match.group(1)
    else:
        return ""

def get_list_of_converted_files():

    '''
    get_list_of_converted_files()

    Returns list of full path of converted text files

    '''
    
    list_text_fnames = []
    
    def append_path(fname):
        return os.path.join(".\converted_pdfs", fname)
    
    logger.info("Reading alltexts.txt, appending converted_pdfs directory name")
    try:
        with open("alltexts.txt", 'r') as fh:
            list_text_fnames = list(fh.read().split('\n'))
    except FileNotFoundError:
        logger.error("alltexts.txt not found")
        sys.exit()

    return list(map(append_path, list_text_fnames))

def format_number_str(s):
    '''
    format_number_str(s)

    Converts number string to float

    Ex: 2,345.00 to 2345.00

    '''
    if s != "":
        return float(s.replace(",", "").replace(" ", ""))
    else:
        return 0  

def month_no_to_name(mnum):
    '''
    month_no_to_name(mnum)

    Returns 3 letter month name from month number

    Ex: 1 -> Jan, 2 -> Feb, 12 -> Dec

    '''
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    return months[mnum-1]


class Payslip:
    '''
    Payslip class

    Contains methods to extract payslip details from a converted text file

    '''
    
    def __init__(self):
        self.pay_period = ""
        self.pay_date = ""
        self.epf_no = ""
        self.uan_no = ""
        self.basic_salary = ""
        self.gross_salary = ""
        self.net_salary = ""
        self.gross_salary_ytd = ""
        self.pf_amount = ""
        self.pf_ytd = ""
        self.income_tax = ""
        self.income_tax_ytd = ""
        self.raw_payslip_text = ""
    
    def read_text(self, fname):
        try:
            with open(fname, 'r') as fh:
                self.raw_payslip_text = fh.read()
                self.pay_period = self.get_pay_period()
                self.pay_date = self.get_pay_date()
                self.epf_no = self.get_epf_number()
                self.uan_no = self.get_uan_number()
                self.basic_salary = self.get_basic_salary()
                self.gross_salary = self.get_gross_sal()
                self.net_salary = self.get_net_sal()
                self.gross_salary_ytd =self.get_gross_sal_ytd()
                self.pf_amount =self.get_pf()
                self.pf_ytd = self.get_pf_ytd()
                self.income_tax = self.get_income_tax()
                self.income_tax_ytd = self.get_income_tax_ytd()
                
        except FileNotFoundError:
            logger.error(f"File not found: {fname}")
            
    def get_pay_period(self):
        match = re.search(r'Pay\sPeriod\s:\s?([\d.]+[\s\-]+[\d.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_pay_date(self):
        match = re.search(r'Pay\sDate\n\n:\s?([\d+.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_epf_number(self):
        match = re.search(r'Emp\sPF\sNumber:\s?([\w\/]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_uan_number(self):
        match = re.search(r'UAN[\n]+:\s?(\d+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""

    def get_basic_salary(self):
        match = re.search(r'Basic\sSalary\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_gross_sal(self):
        match = re.search(r'Total\sGross\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_net_sal(self):
        match = re.search(r'NET\sPAY\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_gross_sal_ytd(self):
        match = re.search(r'YTD\sGROSS\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_pf(self):
        match = re.search(r'Provident\sFund\n+([\d.,]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_pf_ytd(self):
        match = re.search(r'YTD\sEmployee\sPF\n+([\d.,]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_income_tax(self):
        match = re.search(r'Income\sTax\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""
        
    def get_income_tax_ytd(self):
        match = re.search(r'YTD\sTAX\n+([\d,.]+)', self.raw_payslip_text)
        if match:
            return match.group(1)
        else:
            return ""



payslip_details = {
    "pay_period": [],
    "pay_date": [],
    "basic_salary": [],
    "gross_salary": [],
    "net_salary": [],
    "gross_salary_ytd": [],
    "pf_amount": [],
    "pf_ytd": [],
    "income_tax": [],
    "income_tax_ytd": [],
    "epf_no": [],
    "uan_no": [],
}

logger.info("Process started")
logger.info("Converting pdf to text")
convert_pdf_to_text()

list_txt_fnames = get_list_of_converted_files()

pay = Payslip()

logger.info("Saving payslip details from each text file to dictionary")
# Saving payslip details from each text file to dictionary
for fname in list_txt_fnames:
    pay.read_text(fname)
    payslip_details["pay_period"].append(pay.pay_period)
    payslip_details["pay_date"].append(pay.pay_date)
    payslip_details["epf_no"].append(pay.epf_no)
    payslip_details["uan_no"].append(pay.uan_no)
    payslip_details["basic_salary"].append(pay.basic_salary)
    payslip_details["gross_salary"].append(pay.gross_salary)
    payslip_details["net_salary"].append(pay.net_salary)
    payslip_details["gross_salary_ytd"].append(pay.gross_salary_ytd)
    payslip_details["pf_amount"].append(pay.pf_amount)
    payslip_details["pf_ytd"].append(pay.pf_ytd)
    payslip_details["income_tax"].append(pay.income_tax)
    payslip_details["income_tax_ytd"].append(pay.income_tax_ytd)

logger.info("Creating dataframe from dictionary")
# creating dataframe from dictionary
pay_df = pd.DataFrame.from_dict(payslip_details)

logger.info("Formatting columns containing numeric data")
# Formatting columns containing numeric data
# converting object to float
pay_df['basic_salary'] = pay_df['basic_salary'].apply(lambda x: format_number_str(x))
pay_df['net_salary'] = pay_df['net_salary'].apply(lambda x: format_number_str(x))
pay_df['gross_salary'] = pay_df['gross_salary'].apply(lambda x: format_number_str(x))
pay_df['gross_salary_ytd'] = pay_df['gross_salary_ytd'].apply(lambda x: format_number_str(x))
pay_df['pf_amount'] = pay_df['pf_amount'].apply(lambda x: format_number_str(x))
pay_df['pf_ytd'] = pay_df['pf_ytd'].apply(lambda x: format_number_str(x))
pay_df['income_tax'] = pay_df['income_tax'].apply(lambda x: format_number_str(x))
pay_df['income_tax_ytd'] = pay_df['income_tax_ytd'].apply(lambda x: format_number_str(x))

logger.info("Creating Series to hold year and month")
# series to hold month and year
years = pay_df['pay_date'].apply(lambda x: re.sub(r'\d+.\d+.(\d{4})', r'\1',x))
months = pay_df['pay_date'].apply(lambda x: month_no_to_name(int(re.sub(r'\d+.(\d+).\d{4}', r'\1',x))))

logger.info("Appending year and month column to the start")
# appending year and month to the start
pay_df.insert(0,'year',years)
pay_df.insert(1,'months',months)

logger.info("Exporting to Excel")
# exporting to Excel
pay_df.to_excel("payslips.xlsx", index=False)

logger.info("Process completed successfully")