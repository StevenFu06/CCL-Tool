import re
import json
import pandas as pd
import package as pk

with open('treek.json', 'r') as read:
    revk = json.load(read)
with open('treem.json', 'r') as read:
    revm = json.load(read)

print(revk.keys())
print(revm.keys())