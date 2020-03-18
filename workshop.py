import re
import json
import pandas as pd


string = 'Refer to Ill. 144 Assy. D5167890, Ill. 145 Assy. D5168565, Ill. 146 Assy. D5167964 and critical components ' \
         'below '
string2 = 'Refer to Ill. 72 Assy. D2000020483'


def illustration_dict(string):
    string = re.sub(r'\W+', '', string).replace(' ', '')
    if re.findall(r'refertoill.', string, re.IGNORECASE):
        ill_dict = [
            {'num': result[0], 'type': result[1], 'dnum': result[2]}
            for result in re.findall(r'(\d+)(assy|sch)(D\d+)', string, re.IGNORECASE)
        ]
        return ill_dict


def str_to_json(string):
    try:
        string = string.replace("'", '"')
        return json.loads(string)
    except TypeError:
        return []
    except AttributeError:
        return []


def verify(df):
    problems = []
    for idx in df.index:
        row = df.loc[idx]
        if pd.isna(row['pn']):
            problems.append(f'Row {idx}: Missing part number in row')
        if pd.isna(row['fn']):
            problems.append(f'Row {idx}: Missing find number')
        if not pd.isna(row['illustration data']) and not str_to_json(row['illustration data']):
            problems.append(f'Row {idx}: Typo in techincal data')
        if len(str_to_json(row['illustration data'])) != len(str_to_json(row['dnums'])):
            problems.append(f'Row {idx}: Number of documents found does not match number of illustrations found.')
    return problems


df = pd.read_csv('test1.csv')
with open('problems.json', 'w') as write:
    json.dump(verify(df), write, indent=4)
