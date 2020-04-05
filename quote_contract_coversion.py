import pandas as pd
from datetime import datetime, timedelta
quotes = pd.read_csv("./all quotes.csv")
contracts = pd.read_csv("./all_contracts_merged.csv")
TS_name = quotes['TS'].drop_duplicates().values.tolist()
TS_quote_dic = dict()
TS_contract_dic = dict()
for TSname in TS_name:
    TS_quote_dic.update(
        {TSname: quotes.loc[quotes['TS'].str.contains(TSname)]})
    TS_quote_dic[TSname]["Date"] = pd.to_datetime(TS_quote_dic[TSname]["Date"]).dt.to_period(
        "Q")
    TS_contract_dic.update(
        {TSname: contracts.loc[contracts['TS'].str.contains(TSname)]})
    TS_contract_dic[TSname]["Date"] = pd.to_datetime(TS_contract_dic[TSname]["Date"]).dt.to_period(
        "Q")

    try:
        Pn = (Qn_contract_num + Qn_contract_num2*0.7)/len(Qn_quote_list)
    except ZeroDivisionError:
        Pn = 0
    try:
        Pm = (Qm_contract_num + \Qm_contract_num2*0.7)/len(Qm_quote_list)
    except ZeroDivisionError:
        Pm = 0
