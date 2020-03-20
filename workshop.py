import re
import json
import pandas as pd
import package as pk


with open('treek.json', 'r') as read:
    revk = json.load(read)
with open('treem.json', 'r') as read:
    revm = json.load(read)
bomk = pd.read_csv('revk.csv')
bomm = pd.read_csv('revm.csv')


def rec_compare(obja, objb, boma, bomb):
    # Will deal with finding part number
    pnA = [key.split() for key in obja.keys()]
    pnB = [key.split()[1] for key in objb.keys()]
    not_found = [key for key in pnA if key[1] not in pnB]

    # Deals with finding find number matching
    b = [bomb.loc[int(int(b.split()[0])), 'F/N'] for b in objb.keys()]
    again_missing = []
    for item in not_found:
        a = boma.loc[int(item[0]), 'F/N']
        if a not in b:
            again_missing.append(item)
        else:
            obja[f'{item[0]} {pnB[b.index(a)]}'] = obja.pop(' '.join(item))
            print(f'{item} was changed to {item[0]} {pnB[b.index(a)]} using fn')

    # Find matching description
    b = [bomb.loc[int(int(b.split()[0])), 'Description'] for b in objb.keys()]
    again_again_missing = []
    for missing in again_missing:
        a = boma.loc[int(missing[0]), 'Description']
        if a not in b:
            again_again_missing.append(missing)
        else:
            obja[f'{missing[0]} {pnB[b.index(a)]}'] = obja.pop(' '.join(missing))
            print(f'{missing} was changed to {missing[0]} {pnB[b.index(a)]} using desc')

    if again_again_missing:
        print(f'{again_again_missing} were completely not found')
    print()
    # Recursively steps down the dictionary...
    for key in obja:
        objbkeys = list(objb.keys())
        if key.split()[1] in pnB:
            keyb = objbkeys[pnB.index(key.split()[1])]
            if obja[key]:
                rec_compare(obja[key], objb[keyb], boma, bomb)


rec_compare(revk, revm, bomk, bomm)

