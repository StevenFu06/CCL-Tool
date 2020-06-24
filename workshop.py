import re
import pandas as pd
import json

def read_avl(path, skiprow):
    df = pd.read_csv(path, skiprows=skiprow)
    try:
        df['Name']
    except KeyError:
        skiprow += 1
        if skiprow < 10:
            df = read_avl(path, skiprow)
        else:
            raise TypeError('File is in wrong format')
    return df


if __name__ == '__main__':
    df = pd.read_csv('test.csv')
    string = df.loc[0, 'illustration data']
    js = json.loads(string)
    print(js)