import pandas as pd
from pprint import pprint as pprint

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)

df = pd.read_excel("vlan10.xlsx", usecols=['IP ADDRESS', 'NAME', 'MODEL', 'SERIAL NUMBER', 'SW VERSION'])

# removing lines with no switch name with the dropna call
clean_df = df.dropna(subset='NAME')

# empty list to put all devices
device_list = []
# convert into dictionary
dict = clean_df.to_dict('records')
# for i in dict:
#     if i['NAME'] is NaN:
#         continue
#     else:
#         print(i['NAME'])

pprint(dict)
