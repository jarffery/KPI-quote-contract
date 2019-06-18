# -*- coding: utf-8 -*-

import xdrlib, sys, re
import xlrd
from xlrd import xldate_as_tuple
import time
import platform
import tkinter as tk
import pandas as pd
import csv
from datetime import datetime
from datetime import timedelta
# excel using 1900/1/1 as the first day

class KPI(object):
    def __init__(self, file, quote_tab, contract_tab):
        self.file = file
        self.quote_tab = quote_tab
        self.contract_tab = contract_tab
    def open_excel(self):
        try:
            self.data = xlrd.open_workbook(self.file)
            return self.data
        except Exception as e:
            raise NameError (self.file + " cant' be found, please put your summary in the same folder")
    def list(self):
        try:
            self.quote_sheet = self.data.sheet_by_name(self.quote_tab)
            self.contract_sheet = self.data.sheet_by_name(self.contract_tab)
        except Exception as e:
            raise NameError (str(e), "please change your quote tab and contract tab into 'Quote' and 'Contract'!")
        self.quote_list = {}
        self.contract_list = {}
        for n in range(self.quote_sheet.nrows):
            self.quote_row_value = self.quote_sheet.row_values(n)
            if self.quote_row_value:
                for i in range(len(self.quote_row_value)):
                    if type(self.quote_row_value[i]) is str:
                        app = re.findall(r'^NVUS\d+', self.quote_row_value[i])
                        if app:
                            try:
                                self.quote_list.update(
                                    {app[0]: [xlrd.xldate_as_datetime(self.quote_sheet.row_values(n)[0],0), self.quote_sheet.row_values(n)[0]]})
                            except Exception as e:
                                raise KeyError (
                                    str(e), "Please change the data format in to month/day/year.")
        for m in range(self.contract_sheet.nrows):
            self.contract_row_value = self.contract_sheet.row_values(m)
            if self.contract_row_value:
                for p in range(len(self.contract_row_value)):
                    if type(self.contract_row_value[p]) is str:
                        app2 = re.findall(r'^NVUS\d+', self.contract_row_value[p])
                        if app2:
                            try:
                                self.contract_list.update(
                                    {app2[0]: [xlrd.xldate_as_datetime(self.contract_sheet.row_values(m)[0], 0), self.contract_sheet.row_values(m)[0]]})
                            except Exception as e:
                                raise KeyError(
                                    str(e), "Please change the data format in to month/day/year.")
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
        Q = [(datetime(Q_year-1, 3, 1) - timedelta(days=1)), datetime(Q_year-1, 6, 30), datetime(Q_year-1, 9, 30), datetime(Q_year-1, 12, 31)
             , (datetime(Q_year, 3, 1) - timedelta(days=1)), datetime(Q_year, 6, 30), datetime(Q_year, 9, 30), datetime(Q_year, 12, 31)]
        for n in range(4,8):
            if datetime.now() <= Q[n] and datetime.now() > Q[n-1]:
                self.quarter = 4 if n == 4 else n-4
                self.last_quarter = 4 if (self.quarter - 1 == 0) else self.quarter - 1
                self.Qn_days = (Q[n]-Q[n-1]).days
                for key, value in self.contract_list.items():
                    if (value[0] <= Q[n-1]) and (value[0] > Q[n-2]):
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
                    elif (value[0] <= Q[n-2]) and (value[0] > Q[n-3]):
                        self.Qm_days = (Q[n-2]-Q[n-3]).days
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
                    if (value[0] <= Q[n-1]) and (value[0] > Q[n-2]):
                        self.Qn_quote_list.update({key: value[1]})
                    elif (value[0] <= Q[n-2]) and (value[0] > Q[n-3]):
                        self.Qm_quote_list.update({key: value[1]})
                    else:
                        pass
            else:
                pass
        self.Pn = (self.Qn_contract_num + self.Qn_contract_num2*0.7)/len(self.Qn_quote_list)
        self.Pm = (self.Qm_contract_num + self.Qm_contract_num2*0.7)/len(self.Qm_quote_list)

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid(column=0, row=0)
        self.bgColor = '#EEEEEE'
        self.config(bg=self.bgColor, borderwidth=20)
        self.create_widgets()
        self.quoteinfotext.focus()
        # bind shift-enter key to generate quotation
        master.bind_all('<Shift-Return>', lambda e: self.main())

    def paste_quote_text(self, event):
        """
        Binded function to paste clipboard content into Text, after striping
        This is helpful to solve performance issues when a lot of \t are copied from excel
        """
        clipboard = self.clipboard_get()
        self.quoteinfotext.insert('end', clipboard.strip())
        return "break"

    def focus_next(self, event):
        """binded function that switch the focus to the next widget"""
        event.widget.tk_focusNext().focus()
        # return 'break' is a trick to stop the original functionality of the event
        return "break"

    def autoselect(self):
        if self.rRNAremoval_check.get() == True:
            self.library_type.set(True)
        else:
            self.library_type.get()

    def create_widgets(self):
        self.welcome = tk.Label(self, text='KPI calculator: quote and contract transfer rate\n',
                                bg=self.bgColor)
        self.welcome.grid(column=0, row=0, columnspan=4)

        self.quoteinfotext = tk.Text(self, height=2)
        if platform.system() == 'Darwin':
            self.quoteinfotext.bind('<Command-v>', self.paste_quote_text)
        self.quoteinfotext.bind('<Control-v>', self.paste_quote_text)
        self.quoteinfotext.bind("<Tab>", self.focus_next)
        self.quoteinfotext.grid(column=0, row=1, columnspan=4)

        self.label1 = tk.Label(
            self, text='please input the name of your summary file', bg=self.bgColor)
        self.label1.grid(column=0, row=2, columnspan=4)

        self.run = tk.Button(self, text='run',
                             command=self.main, highlightbackground=self.bgColor)
        self.run.grid(column=3, row=11, columnspan=1)

        #self.run = tk.Button(self, text='submit contract',
        #                    command=self.main, highlightbackground=self.bgColor)
        #self.run.grid(column=2, row=11, columnspan=1)

        self.clear = tk.Button(
            self, text='Clear All', command=self.clearall, highlightbackground=self.bgColor)
        self.clear.grid(column=0, row=11, columnspan=1)

        self.errorLabel = tk.Label(self, bg=self.bgColor, fg='red')
        self.errorLabel.grid(column=0, row=12, columnspan=8)

    def main(self):
        try:
            file = str(self.quoteinfotext.get('1.0', 'end').strip() + ".xlsx")
            quote_tab = "Quote"
            contract_tab = "Contract"
            a = KPI(file, quote_tab, contract_tab)
            a.open_excel()
            a.list()
            a.KPI_calculate()
            #generate the csv
            with open('Quarter_report.csv', 'w', newline="") as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(["quote", "quote date", "contract date"])
                for key, value in a.Qn_quote_list.items():
                    value2 = None
                    if a.contract_list.get(key):
                        value2 = a.contract_list[key][1]
                    writer.writerow([key, value, value2])
            with open('Quarter_report2.csv', 'w', newline="") as csv_file2:
                writer = csv.writer(csv_file2)
                writer.writerow(["quote", "quote date", "contract date"])
                for key, value in a.Qm_quote_list.items():
                    value2 = None
                    if a.contract_list.get(key):
                        value2 = a.contract_list[key][1]
                    writer.writerow([key, value, value2])
            with open('contract_withuot_quote_date.csv', 'w', newline="") as csv_file3:
                writer = csv.writer(csv_file3)
                writer.writerow(["contract", "date"])
                for key, value in a.error_list.items():
                    writer.writerow([key, value])
            with open('report.txt', 'w') as f:
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
                "please check the output file: report.txt, Quarter_report.csv, Quarter_report2.csv, contract_withuot_quote_date.csv")
            self.errorLabel.config(text= output)
        except (NameError, Exception) as e:
            self.errorLabel.config(text=str(e))
            raise

    def clearall(self):
        self.quoteinfotext.delete('1.0', 'end')
        self.pricetext1.delete('1.0', 'end')
        self.pricetext2.delete('1.0', 'end')
        self.errorLabel.config(text='')

root = tk.Tk()
app = Application(master=root)
app.mainloop()



      
