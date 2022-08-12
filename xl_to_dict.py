"""
converts network overview spreadsheet to either json or yaml for use as a
device list inventory to work with netmiko
"""

import json
import yaml
import pandas as pd
import credentials


def read_to_pandas_df():
    """
    read excel file from directory and put in a panda dataframe for manipulation
    :return: device_list
    """
    pd.set_option("display.max_columns", None)
    pd.set_option("display.max_rows", None)
    pd.set_option("display.width", 1000)

    data_frame = pd.read_excel(
        "vlan10.xlsx",
        usecols=["IP ADDRESS", "NAME", "MODEL", "SERIAL NUMBER", "SW VERSION"],
        sheet_name="Sheet1",
    )

    # removing lines with no switch name with the dropna call
    clean_df = data_frame.dropna(subset=["NAME", "MODEL", "SW VERSION"])

    # empty list to put all devices
    device_list = []

    # convert into dictionary
    device_dict = clean_df.to_dict("records")

    for i in device_dict:
        if i["NAME"] == " " or i["MODEL"] == "Future" or i["MODEL"] == "PA-220":
            continue
        hosts = {"hostname": f'{i["NAME"]}', "host": f'{i["IP ADDRESS"]}'}
        device_list.append(hosts)
    return device_list


def convert_to_json():
    """
    # convert to json format for storing
    """
    device_list = read_to_pandas_df()
    device_list_json = json.dumps(
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
    with open("device_list.json", "w", encoding="utf8") as j:
        j.write(device_list_json)


def convert_to_yaml():
    """
    # convert to yaml
    """
    device_list = read_to_pandas_df()
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
    with open("device_list.yaml", "w", encoding="utf8") as j:
        j.write(device_list_yaml)


def main():
    """
    convert to either json or yaml or both
    :return:
    """
    convert_to_yaml()
    convert_to_json()


if __name__ == "__main__":
    main()
