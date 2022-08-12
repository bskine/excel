from pprint import pprint
import json
import yaml
import pandas as pd
import credentials

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.width", 1000)

df = pd.read_excel(
    "vlan10.xlsx",
    usecols=["IP ADDRESS", "NAME", "MODEL", "SERIAL NUMBER", "SW VERSION"],
    sheet_name="Sheet1",
)

# removing lines with no switch name with the dropna call
clean_df = df.dropna(subset=["NAME", "MODEL", "SW VERSION"])

# empty list to put all devices
device_list = []
# convert into dictionary
device_dict = clean_df.to_dict("records")

for i in device_dict:
    if i["NAME"] == " " or i["MODEL"] == "Future" or i["MODEL"] == "PA-220":
        continue
    else:
        hosts = {"hostname": f'{i["NAME"]}', "host": f'{i["IP ADDRESS"]}'}
        device_list.append(hosts)

# # convert to json format for storing
# device_list_json = json.dumps(device_list, indent=2)
# with open("device_list.json", "w") as j:
#     j.write(device_list_json)
# # pprint(device_dict)

# convert to yaml???
device_list_yaml = yaml.dump(
    {
        "common_vars": {
            "username": credentials.username,
            "password": credentials.password,
            "device_type": "cisco_ios",
        },
        "hosts": device_list,
    },
    indent=4,
)
with open("device_list.yaml", "w") as j:
    j.write(device_list_yaml)
