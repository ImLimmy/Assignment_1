import pandas as pd
import json
from datetime import datetime

# converting the datetime into yyyy/mm/dd
class DateTimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.strftime('%Y/%m/%d')
        return super().default(obj)

# read the data from the Excel file into a dataframe
df = pd.read_excel("data.xlsx", keep_default_na=False, dtype=object)

output = {}

# Iterate over each row in the dataframe
for ind in df.index:
    if df['ID'][ind] != "":
        ID = df['ID'][ind]
        
        if ID not in output:
            output[ID] = {
                'name': df['Name'][ind],
                'exposure': []
            }
        
        if df['Date'][ind] != "" and df['Deed'][ind] != "":
            exposure_data = {
                'Date': df['Date'][ind],
                'Deed': df['Deed'][ind]
            }
            output[ID]['exposure'].append(exposure_data)

# sorting the exposure date within each ID in ascending order
for key, value in output.items():
    value['exposure'] = sorted(value['exposure'], key=lambda x: x['Date'])

# specify the file path for the JSON file
json_file_path = 'output.json'

# write each ID as a separate JSON object in the file
with open(json_file_path, 'w') as json_file:
    for key, value in output.items():
        # Convert the output dictionary to JSON and use the DateTimeEncoder to handle datetime objects
        json_output = json.dumps({'ID': key, 'Name': value['name'], 'Exposure': value['exposure']}, cls=DateTimeEncoder, indent=4)
        json_file.write(json_output + '\n\n')

print(f"Data has been imported to {json_file_path}")