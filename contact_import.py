import pandas as pd
from csv2vcard import csv2vcard

data = pd.read_excel("vcard.xlsx")
data.drop_duplicates(subset="Phone Number", inplace=True, keep="last")

with open("contacts.csv", "w", encoding="UTF-8") as file:
    file.write("last_name,first_name,org,title,phone,email,website,street,city,p_code,country" + "\n")
    for i in range(len(data)):
        file.write(f",{str(data.iloc[i]['Name']).title()},,,{data.iloc[i]['Phone Number']},,,,,," + "\n")

csv2vcard.csv2vcard("contacts.csv", ",")
