import pandas as pd

quotes = pd.read_csv("./all quotes.csv")
contracts = pd.read_csv("./all_contracts_merged.csv")
TS_name = quotes['TS'].drop_duplicates().values.tolist()
TS_quote_dic = dict()
TS_contract_dic = dict()
for TSname in TS_name:
    TS_quote_dic.update(
        {TSname: quotes.loc[quotes['TS'].str.contains(TSname)]})
    TS_contract_dic.update(
        {TSname: contracts.loc[contracts['TS'].str.contains(TSname)]})
