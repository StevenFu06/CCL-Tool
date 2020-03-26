import re
import json
import pandas as pd
import package as pk
from fuzzywuzzy import fuzz


with open('treek.json', 'r') as read:
    revk = json.load(read)
bomk = pd.read_csv('revk.csv')

with open('treeg.json', 'r') as read:
    revg = json.load(read)
bomg = pd.read_csv('revg.csv')

total_iter = 0


def type_split(a, b):
    desca = a[1]
    descb = b[1]
    if '*' in desca and '*' in descb:
        if descb.split('*')[0] == descb.split('*')[0]:
            return True
    return False


def is_updated(a, b):
    for item in b:
        if a[0] == item[0] and (fuzz.ratio(a[1], item[1]) >= 50 or type_split(a, item)):
            print(fuzz.ratio(a[1], item[1]))
            return item
    return False


def rec_compare(obja, objb, boma, bomb):
    # Will deal with finding part number
    global total_iter
    total_iter += 1
    pnA = [key.split() for key in obja.keys()]
    pnB = [key.split()[1] for key in objb.keys()]
    not_found = [key for key in pnA if key[1] not in pnB]

    # Deals with finding find number matching
    b = [[bomb.loc[int(int(b.split()[0])), "F/N"], bomb.loc[int(int(b.split()[0])), "Description"]]
         for b in objb.keys()]
    again_missing = []
    for item in not_found:
        a = [boma.loc[int(item[0]), "F/N"], boma.loc[int(item[0]), "Description"]]
        new_index = is_updated(a, b)
        if new_index is False:
            again_missing.append(item)
        else:
            obja[f'{item[0]} {pnB[b.index(new_index)]}'] = obja.pop(' '.join(item))
            print(f'{item} was changed to {item[0]} {pnB[b.index(new_index)]} using fn')

    if again_missing:
        print(f'{again_missing} were completely not found')
    # Recursively steps down the dictionary...
    for key in obja:
        objbkeys = list(objb.keys())
        if key.split()[1] in pnB:
            keyb = objbkeys[pnB.index(key.split()[1])]
            if obja[key]:
                rec_compare(obja[key], objb[keyb], boma, bomb)


rec_compare(revg, revk, bomg, bomk)
print(total_iter)
