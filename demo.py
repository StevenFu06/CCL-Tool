from ccl import CCL
from filehandler import DocumentCollector, Illustration
from compare import Tracker, Rearrange, Bom

import os
import pandas as pd
import json

import pickle


########################################################################################################################
# Features / Functions
########################################################################################################################
def bom_comparison(avl_bom_old, tree_old, avl_bom_new, tree_new):
    tracker = Tracker()
    bom_old = Bom(avl_bom_old, tree_old)
    bom_new = Bom(avl_bom_new, tree_new)
    Rearrange(bom_old, bom_new, tracker)
    print(tracker.not_found)


def update_ccl(ccl, avl_bom_old, avl_bom_new, save_loc):
    ccl = CCL(ccl, avl_bom_old)
    ccl.update(avl_bom_new)
    ccl.save(save_loc)


def collect_documents(ccl, save_loc, check_locations, username, password, processes):
    collector = DocumentCollector(username, password, ccl, save_loc, processes=processes, headless=False)
    collector.collect_documents(check_locations)


def collect_illustrations(ccl_filtered, save_loc, ccl_data, processes):
    drawings = Illustration(ccl, save_loc, processes=processes)
    drawings.filtered = ccl_filtered
    drawings.get_illustrations(ccl_dir=ccl_data)

def insert_illustration(ccl, ill_loc, ill_num):
    drawings = Illustration(ccl, ill_loc)
    drawings.shift_up_ill(ill_num)

def delete_illustration(ccl, ill_loc, ill_num):
    drawings = Illustration(ccl, ill_loc)
    drawings.shift_down_ill(ill_num)


if __name__ == '__main__':
    import pandas as pd

    ccl = 'rev c bugatti.docx'
    save_loc = 'ccl docs'
    check_locations = None
    username = 'Steven.Fu'
    password = pickle.load(open('password.pkl', 'rb'))
    processes = 2

    wombat_ccl_filtered = pd.read_csv('filter.csv', index_col = 0)
    wombat_docs = 'Wombat Jorb CCL Documents'

    # collect_documents(ccl, save_loc, check_locations, username, password, processes)
    collect_illustrations(wombat_ccl_filtered, 'illustrations', wombat_docs, processes)
