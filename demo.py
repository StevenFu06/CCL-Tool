from ccl import CCL
from filehandler import DocumentCollector, Illustration


########################################################################################################################
# Features / Functions
########################################################################################################################
def update_ccl(ccl, avl_bom_old, avl_bom_new):
    ccl = CCL(ccl, avl_bom_old)
    ccl.update(avl_bom_new)

def collect_documents(ccl, save_loc, check_locations, username, password, processes):
    collector = DocumentCollector(username, password, ccl, save_loc, processes=processes)
    collector.collect_documents(check_locations)

def update_documents(ccl, save_loc, username, password):
    pass

def collect_illustrations(ccl, save_loc, ccl_data, processes):
    drawings = Illustration(ccl, save_loc, processes=processes)
    drawings.get_illustrations(ccl_dir=ccl_data)

def edit_ccl_illustration(ccl, ill_loc):
    pass

def insert_illustration(ccl, save_loc):
    pass

def delete_illustration(ccl, save_loc):
    pass

########################################################################################################################
# Inputs / Documents
########################################################################################################################

########################################################################################################################
# Demo Code below
########################################################################################################################
