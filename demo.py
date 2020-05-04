from ccl import CCL
from filehandler import DocumentCollector, Illustration
from compare import Tracker, Rearrange, Bom

import os
import pandas as pd
import json


########################################################################################################################
# Features / Functions
########################################################################################################################
def bom_comparison(avl_bom_old, tree_old, avl_bom_new, tree_new):
    tracker = Tracker()
    bom_old = Bom(avl_bom_old, tree_old)
    bom_new = Bom(avl_bom_new, tree_new)
    Rearrange(bom_old, bom_new, tracker)


def update_ccl(ccl, avl_bom_old, avl_bom_new, save_loc):
    ccl = CCL(ccl, avl_bom_old)
    ccl.update(avl_bom_new)
    ccl.save(save_loc)


def collect_documents(ccl, save_loc, check_locations, username, password, processes):
    collector = DocumentCollector(username, password, ccl, save_loc, processes=processes)
    collector.collect_documents(check_locations)


def collect_illustrations(ccl, save_loc, ccl_data, processes):
    drawings = Illustration(ccl, save_loc, processes=processes)
    drawings.get_illustrations(ccl_dir=ccl_data)


def update_ccl_illustration(ccl, ill_loc, updated_ccl):
    drawings = Illustration(ccl, ill_loc)
    drawings.update_ccl(updated_ccl)


def insert_illustration(ccl, ill_loc, ill_num):
    drawings = Illustration(ccl, ill_loc)
    drawings.shift_up_ill(ill_num)


def delete_illustration(ccl, ill_loc, ill_num):
    drawings = Illustration(ccl, ill_loc)
    drawings.shift_down_ill(ill_num)


########################################################################################################################
# Inputs / Documents
########################################################################################################################

# BOM Comparison
wombat_rev_g = pd.read_csv(os.path.join('Demo', 'Bom Comparison', 'wombat revg.csv'))
wombat_tree_g = os.path.join('Demo', 'Bom Comparison', 'treeg.json')
with open(wombat_tree_g, 'r') as read:
    wombat_tree_g = json.load(read)

wombat_rev_n = pd.read_csv(os.path.join('Demo', 'Bom Comparison', 'wombat revn.csv'))
wombat_tree_n = os.path.join('Demo', 'Bom Comparison', 'treen.json')
with open(wombat_tree_n, 'r') as read:
    wombat_tree_n = json.load(read)

# Update CCL
bugatti_ccl = os.path.join('Demo', 'Update CCL', 'rev a bugatti.docx')
updated_ccl = os.path.join('Demo', 'Update CCL', 'rev c bugatti.docx')
bugatti_rev_a = os.path.join('Demo', 'Update CCL', 'reva.csv')
bugatti_rev_c = os.path.join('Demo', 'Update CCL', 'revc.csv')

# Collect Documents
"""Move over to laptop for collect documents demo"""

########################################################################################################################
# Demo Code below
########################################################################################################################

# bom_comparison(wombat_rev_g, wombat_tree_g, wombat_rev_n, wombat_tree_n)
# update_ccl(bugatti_ccl, bugatti_rev_a, bugatti_rev_c, updated_ccl)