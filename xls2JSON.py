import pandas as pd
from dateutil.parser import parse
import json

def convert_date(date):
    if pd.isnull(date):
        return None
    if isinstance(date, pd.Timestamp):
        return date.isoformat()
    try:
        return pd.to_datetime(date).isoformat()
    except ValueError:
        try:
            return parse(date).isoformat()
        except (ValueError, TypeError):
            print(f"Couldn't parse date: {date}")
            return None

# Load the Excel file into a pandas DataFrame
df = pd.read_excel('XLS2JSON.xlsx', engine='openpyxl')

# Initialize an empty list to hold the dictionaries
data = []

# Loop through each row in the DataFrame and append the corresponding dictionary to the list
for index, row in df.iterrows():
    payload = {
        "oid": row['oid'],
        "sourceName": row['sourceName'],
        "externalReferenceID": row['externalReferenceID'],
        "legalEntityCode": row['legalEntityCode'],
        "clientID": row['clientID'],
        "dealNumber": row['dealNumber'],
        "bondType": row['bondType'],
        "instrumentShortName": row['instrumentShortName'],
        "originalAcqDate": convert_date(row['originalAcqDate']),
        "tradeDate": convert_date(row['tradeDate']),
        "settlementDate": convert_date(row['settlementDate']),
        "quantity": row['quantity'],
        "dirtyPrice": row['dirtyPrice'],
        "cleanPrice": row['cleanPrice'],
        "residualFactor": row['residualFactor'],
        "fxRate": row['fxRate'],
        "lastInterestDate": convert_date(row['lastInterestDate']),
        "isrAmount": row['isrAmount'],
        "amount": row['amount'],
        "amountCCYIsoCode": row['amountCCYIsoCode'],
        "direction": row['direction'],
        "days": row['days'],
        "nominalInterest": row['nominalInterest'],
        "realInterest": row['realInterest']
    }
    data.append({
        "id": row['id'],
        "action": row['action'],
        "entity": row['entity'],
        "payload": payload
    })

# Write the list of dictionaries to a json file
with open('output.json', 'w') as outfile:
    json.dump(data, outfile, ensure_ascii=False, indent=4)
