import pandas as pd
import re

ledger_df = pd.read_csv('Ledger23-24.csv', skiprows=3)
sales_df = pd.read_excel('BDM-data.xlsx', sheet_name='SalesData',skiprows=12, engine='openpyxl')
 
# Removing completely empty rows
ledger_df = ledger_df.dropna(how="all") 

# Keeping only rows with valid account names
ledger_df = ledger_df[ledger_df["A/c Name"].notna()]

# Filling down merged values
ledger_df["A/c Name"].fillna(method="ffill", inplace=True)  

# Function to extract Bill Nos
def extract_bills(row):
    try:
        match = re.search(r'BILLNO\s*:\s*(.*)', str(row))
        if match:
            bills = match.group(1).replace(' ', '').split(',')
            return bills
    except:
        return []
    
records = []

# Creating new rows from ledger: one per bill
for idx, row in ledger_df.iterrows():
    # Checking if the A/C Name contains 'BILLNO' & Debit is not Nan ( as those are the incoming payments )
    if str(row['A/c Name']).__contains__('BILLNO') and pd.notna(row['Debit']):
        buyer = str(row['A/c Name']).split('\n')[0] 
        bill_cell = row['A/c Name']
        bills = extract_bills(bill_cell)
        for bill in bills:
            records.append({
                'BILL NO': bill,
                'PARTY': buyer.replace(' ', ''),
                'PAYMENT DATE': row['Date'],
                'Amount': row['Debit']
            })

ledger_cleaned = pd.DataFrame(records)

# Sorting by payment date, and dropping duplicate payment record for smae Bill No. 
ledger_latest = (
    ledger_cleaned
    .sort_values("PAYMENT DATE")
    .drop_duplicates(subset="BILL NO", keep="last")
)
ledger_cleaned.to_excel("ledger23-24-cleaned.xlsx", index=False)

sales_df['BILL NO'] = sales_df['BILL NO'].astype(str).str.replace(' ', '')

# Merging the ledger records against sales by Bill No
final_df = pd.merge(sales_df, ledger_latest, on='BILL NO', how='left')

# Removing the rows where payment date is NaN
final_df = final_df.dropna(subset=['PAYMENT DATE']) 

final_df.to_excel("final_report.xlsx", index=False)
