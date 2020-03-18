import json
import pandas as pd


def str_to_json(string):
    """Converts the string to a dict/ list"""

    try:
        # Json.loads needs " instead of '
        string = string.replace("'", '"')
        return json.loads(string)
    except TypeError:
        return []
    except AttributeError:
        return []


def verify_filtered(df):
    """Verification for filtered.csv

    Checks:
        - Missing part number
        - Missing find number
        - Improper parsing of techincal data

    return: a list of all the problems in the filtered.csv
    """
    problems = []
    for idx in df.index:
        row = df.loc[idx]
        if pd.isna(row['pn']):  # Missing part number check
            problems.append(f'Row {idx}: Missing part number in row')
        if pd.isna(row['fn']):  # Missing find number check
            problems.append(f'Row {idx}: Missing find number')

        # Checks for improper parsing of techincal data
        if not pd.isna(row['illustration data']) and not str_to_json(row['illustration data']):
            problems.append(f'Row {idx}: Typo in techincal data')
        if len(str_to_json(row['illustration data'])) != len(str_to_json(row['dnums'])):
            problems.append(f'Row {idx}: Number of documents found does not match number of illustrations found.')
    return problems
