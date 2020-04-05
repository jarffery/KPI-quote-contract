import pandas as pd

#set the Quarter, suold be yearQn
Quarter = "2020Q1"

#read CSV file: quote and contract
quotes = pd.read_csv("./all quotes.csv")
contracts = pd.read_csv("./all_contracts_merged.csv")

#get the TS name list
TS_name = quotes['TS'].drop_duplicates().values.tolist()

#set up the blank dict
TS_quote_dic = dict()
TS_contract_dic = dict()
Quarter_dic = dict()

#set up a new black dataframe
df_summary = pd.DataFrame(columns=['TS', 'Total Quote', 'Total Contract', 'Contract in this quarter','Conversion Rate'])

#get all detail for each TS
for TSname in TS_name:
    TS_quote_dic.update(
        {TSname: quotes.loc[quotes['TS'].str.contains(TSname)]})
    TS_quote_dic[TSname]["Date"] = pd.to_datetime(TS_quote_dic[TSname]["Date"]).dt.to_period("Q")
    TS_contract_dic.update(
        {TSname: contracts.loc[contracts['TS'].str.contains(TSname)]})
    TS_contract_dic[TSname]["Contract number generation date"] = pd.to_datetime(
        TS_contract_dic[TSname]["Contract number generation date"]).dt.to_period("Q")
    Contract_in_quarter = TS_contract_dic[TSname][TS_contract_dic[TSname]["Contract number generation date"] == Quarter]
    Quote_in_quarter = TS_quote_dic[TSname][TS_quote_dic[TSname]["Date"] == Quarter]
    Quote_in_quarter_num = Quote_in_quarter.shape[0]
    total_contract_num = Contract_in_quarter.shape[0]
    Contract_in_quarter_num = Contract_in_quarter["Quotation No."].isin(Quote_in_quarter["Quotation No."]).shape[0]
    try:
        conversion_rate = (Contract_in_quarter_num + (total_contract_num - Contract_in_quarter_num) * 0.7)/Quote_in_quarter_num
    except ZeroDivisionError:
        conversion_rate = 0
    Quarter_dic.update({TSname: {
        "Quote": Quote_in_quarter, 
        "Contract": Contract_in_quarter, 
        "Total_contract_num": total_contract_num,
        "Contract_in_quarter_num": Contract_in_quarter_num,
        "Conversion_rate": conversion_rate,
        }})
    df_summary = df_summary.append(pd.DataFrame({
        'TS': [TSname], 
        'Total Quote': [Quote_in_quarter_num],
        'Total Contract': [total_contract_num],
        'Contract in this quarter': [Contract_in_quarter_num],
        'Conversion Rate': [conversion_rate],
        }), ignore_index=True)

df_summary.to_csv(r'./conversion.csv', index=False)
