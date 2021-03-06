# -*- coding: utf-8 -*-

import xdrlib, sys, re
import xlrd
from xlrd import xldate_as_tuple
import time
import csv
from datetime import datetime
from datetime import timedelta
import os
import shutil
# excel using 1900/1/1 as the first day

class KPI(object):
    def __init__(self, quote, contract):
        self.quote_file = quote
        self.contract_file = contract
        #self.quarter = int(quarter_time)
    def open_excel(self):
        try:
            self.quote = xlrd.open_workbook(r''.join(self.quote_file))
            self.contract = xlrd.open_workbook(r''.join(self.contract_file))
            return self.quote, self.contract
        except Exception as e:
            raise NameError (self.quote + "or" + self.quote + " cant' be found, please check your dir and try again")
    def open_list(self):
        try:
            self.quote_sheet = self.quote.sheet_by_name(self.quote_tab)
            self.contract_sheet = self.data.sheet_by_name(self.contract_tab)
        except Exception as e:
            raise NameError ("please change your quote tab and contract tab into 'Quote' and 'Contract'!")
        self.quote_list = {}
        self.contract_list = {}
        self.error_quote_list = []
        self.error_contract_list = []
        for n in range(self.quote_sheet.nrows):
            self.quote_number = self.quote_sheet.cell(n,1).value
            if self.quote_number:
                app = re.findall(r'^NVUS\d+', self.quote_number)
                if app and (type(self.quote_sheet.row_values(n)[0]) == float):
                    try:
                        self.quote_list.update(
                            {app[0]: [xlrd.xldate_as_datetime(self.quote_sheet.row_values(n)[0],0), self.quote_sheet.row_values(n)[0]]})
                    except Exception as e:
                        raise KeyError (str(app) + ": " + str(e))
                else:
                    if app:
                        self.error_quote_list.append(app[0])
                    else:
                        pass
            else:
                pass
        for m in range(self.contract_sheet.nrows):
            self.contract_quote = self.contract_sheet.cell(m, 3).value
            if self.contract_quote:
                app2 = re.findall(r'^NVUS\d+', self.contract_quote)
                if app2 and type(self.contract_sheet.row_values(m)[0]) == float:
                    try:
                        self.contract_list.update(
                            {app2[0]: [xlrd.xldate_as_datetime(self.contract_sheet.row_values(m)[0], 0), self.contract_sheet.row_values(m)[0]]})
                    except Exception as e:
                        raise KeyError(str(app) + ": " + str(e))
                else:
                    if app2:
                        self.error_contract_list.append(app2[0])
                    else:
                        pass
            else:
                pass
        if self.quote_list == {} or self.contract_list == {}:
            raise KeyError(
                "please change the summary excel's column structure")
    def KPI_calculate(self):
        self.Qn_contract_list = {}
        self.Qm_contract_list = {}
        self.Qn_quote_list = {}
        self.Qm_quote_list = {}
        self.error_list = {}
        self.error_count = 0
        self.Qn_contract_num = 0
        self.Qn_contract_num2 = 0
        self.Qm_contract_num = 0
        self.Qm_contract_num2 = 0
        Q_year = datetime.now().timetuple().tm_year
        Q = [datetime(Q_year-1, 3, 31), datetime(Q_year-1, 6, 30), datetime(Q_year-1, 9, 30), datetime(Q_year-1, 12, 31)
             , datetime(Q_year, 3, 31), datetime(Q_year, 6, 30), datetime(Q_year, 9, 30), datetime(Q_year, 12, 31)]
        for n in range(3,8):
            if datetime.now() <= (Q[n] + timedelta(days = 20)) and datetime.now() > (Q[n-1] + timedelta(days = 20)):
                self.quarter = n + 1 if n == 3 else n-3
                self.last_quarter = 4 if (self.quarter - 1 == 0) else (self.quarter - 1)
                self.Qn_days = (Q[n]-Q[n-1]).days
                for key, value in self.contract_list.items():
                    if (value[0] <= Q[n]) and (value[0] > Q[n-1]):
                        self.Qn_contract_list.update({key: value[1]})
                        if key in self.quote_list.keys():
                            self.Qn_contract_list[key] = self.contract_list[key][1] - self.quote_list[key][1]
                            if self.Qn_contract_list[key] >= self.Qn_days:
                                self.Qn_contract_num2 += 1
                            else:
                                self.Qn_contract_num += 1
                        else:
                            self.error_count += 1
                            self.error_list.update({key: value[1]})
                    elif (value[0] <= Q[n-1]) and (value[0] > Q[n-2]):
                        self.Qm_days = (Q[n-1]-Q[n-2]).days
                        self.Qm_contract_list.update({key: value[1]})
                        if key in self.quote_list.keys():
                            self.Qm_contract_list[key] = self.contract_list[key][1] - self.quote_list[key][1]
                            if self.Qm_contract_list[key] >= self.Qm_days:
                                self.Qm_contract_num2 += 1
                            else:
                                self.Qm_contract_num += 1
                        else:
                            self.error_count += 1
                            self.error_list.update({key: value[1]})
                    else:
                        pass
                for key, value in self.quote_list.items():
                    if (value[0] <= Q[n]) and (value[0] > Q[n-1]):
                        self.Qn_quote_list.update({key: value[1]})
                    elif (value[0] <= Q[n-1]) and (value[0] > Q[n-2]):
                        self.Qm_quote_list.update({key: value[1]})
                    else:
                        pass
            else:
                pass
        try:
            self.Pn = (self.Qn_contract_num + self.Qn_contract_num2*0.7)/len(self.Qn_quote_list)
        except ZeroDivisionError:
            self.Pn = 0
        try:
            self.Pm = (self.Qm_contract_num + self.Qm_contract_num2*0.7)/len(self.Qm_quote_list)
        except ZeroDivisionError:
            self.Pm = 0
                     
    def main(self):
        try:
            file = str(self.quoteinfotext.get('1.0', 'end').strip().replace("\\", "/").replace("\1", "/1"))
            TS_name = str(self.quoteinfotext2.get('1.0', 'end').strip())
            #Quarter_time = int(self.quoteinfotext2.get('1.0', 'end'))
            a = KPI(file)
            a.open_excel()
            a.open_list()
        except (NameError, Exception) as e:
            raise self.errorLabel.config(text=str(e))
        try:
            #dir_name = os.path.split(file)[0]
            dir_name = os.path.split(file)[0]
            dir_name = dir_name + "/report" + "-" + TS_name
            os.mkdir(dir_name)
        except FileExistsError as e:
            raise self.errorLabel.config(text=("Folder " + dir_name + " already exists"))
        #calculate KPI and make the output dir
        if (not bool(a.error_quote_list)) and (not bool(a.error_contract_list)):
            try:
                a.KPI_calculate()
            except Exception as e:
                raise self.errorLabel.config(text=str(e))
            #write the file into the dir
            try:
                with open((dir_name + '/' + 'Quarter_report.csv'), 'w', newline="") as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerow(["quote", "quote date", "contract date"])
                    for key, value in a.Qn_quote_list.items():
                        value2 = None
                        if a.contract_list.get(key):
                            value2 = a.contract_list[key][1]
                        writer.writerow([key, value, value2])
                with open((dir_name + '/' + 'Quarter_report2.csv'), 'w', newline="") as csv_file2:
                    writer = csv.writer(csv_file2)
                    writer.writerow(["quote", "quote date", "contract date"])
                    for key, value in a.Qm_quote_list.items():
                        value2 = None
                        if a.contract_list.get(key):
                            value2 = a.contract_list[key][1]
                        writer.writerow([key, value, value2])
                with open((dir_name + '/' + 'contract_withuot_quote_date.csv'), 'w', newline="") as csv_file3:
                    writer = csv.writer(csv_file3)
                    writer.writerow(["contract", "date"])
                    for key, value in a.error_list.items():
                        writer.writerow([key, value])
                with open((dir_name + '/' + 'report-' + TS_name + '.txt'), 'w') as f:
                    f.write( 
                    "the output is last two quarter's details(based on the current time): \n\n\n"
                    "Quarter " + str(a.quarter) + '\n'
                    "the total number of quote is: " + str(len(a.Qn_quote_list)) + '\n'
                    "the total number of contract is: " + str(len(a.Qn_contract_list)) + '\n'
                    "the total number of contract within this quarter is: " + str(a.Qn_contract_num) + '\n'
                    "the total number of contract outside of this quarter is: " + str(a.Qn_contract_num2) + '\n'
                    "transfer rate is: " + str(a.Pn) + '\n\n'
                    "Quarter " + str(a.last_quarter) + '\n'   
                    "the total number of quote is: " + str(len(a.Qm_quote_list)) + '\n'
                    "the total number of contract is: " + str(len(a.Qm_contract_list)) + '\n'
                    "the total number of contract within this quarter is: " + str(a.Qm_contract_num) + '\n'
                    "the total number of contract outside of this quarter is: " + str(a.Qm_contract_num2) + '\n'
                    "transfer rate is: " + str(a.Pm) + '\n\n'
                    "please check the Quarter_report.csv for details")
                output = str(
                    "please check the output file in the report forlder: report.txt, Quarter_report.csv, Quarter_report2.csv, contract_withuot_quote_date.csv")
                self.errorLabel.config(text= output)
            except (NameError, Exception) as e:
                self.errorLabel.config(text=str(e))
                raise
        else:
            with open((dir_name + '/' + 'error.csv'), 'w', newline="") as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(["error date in quote"])
                for key in a.error_quote_list:
                    writer.writerow([key])
                writer.writerow(["error date in contract"])
                for key in a.error_contract_list:
                    writer.writerow([key])
            self.errorLabel.config(text=str("please find the qoute with wrong date format in the report folder, change your summary and try again"))
            raise




      
