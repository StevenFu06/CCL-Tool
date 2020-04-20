import re
import json
import pandas as pd
import package as pk
from fuzzywuzzy import fuzz
from compare import Bom
from docx.api import Document


with open('treec.json', 'r') as read:
    revc = json.load(read)
bomc = pd.read_csv('revc.csv')

with open('treea.json', 'r') as read:
    reva = json.load(read)
boma = pd.read_csv('reva.csv')


def difference(a, b):
    """Find difference between a and b (a intersect not b)

    Effectively if a = [1,2,30,4,5] b = [24,5,1,2,3,12]
    will return [4, 30]

    :return [[idx, pn],[idx, pn],[idx, pn]...]
    """
    pa = [pn.split()[1] for pn in a]
    pb = [pn.split()[1] for pn in b]
    difference_list = []
    for pn in range(len(pa)):

        if pa[pn] not in pb:
            index_a = int(a[pn].split()[0])
            difference_list.append([index_a, pa[pn]])

    return difference_list

def intersect(a, b):
    pa = [pn.split()[1] for pn in a]
    pb = [pn.split()[1] for pn in b]
    intersect_list = []
    for pn in range(len(pa)):
        if pa[pn] in pb:
            index_a = int(a[pn].split()[0])
            index_b = int(b[pb.index(pa[pn])].split()[0])

            b.pop(pb.index(pa[pn]))
            pb.pop(pb.index(pa[pn]))

            intersect_list.append([index_a, index_b])
    return intersect_list


def ismatch(old_tuple, new_tuple, threshold_full=50, threshold_partial=80):
    old_fn, old_desc = old_tuple
    new_fn, new_desc = new_tuple
    match_pct = fuzz.ratio(old_desc, new_desc)
    old_type, new_type = old_desc.split('*')[0], new_desc.split('*')[0]

    if old_fn == new_fn and (match_pct >= threshold_full or old_type == new_type):
        return 'full'
    elif match_pct >= threshold_partial:
        return 'partial'
    return False


def is_updated(old_tuple, exclusive_new: list, threshold_full=50, threshold_partial=80):
    for tuple_new in exclusive_new:
        match = ismatch(old_tuple, tuple_new, threshold_full, threshold_partial)
        if match == 'full' or match == 'partial':
            return match, tuple_new
    return False, False


not_found = []
def recursive_compare(old_parent, old_bom, new_parent, new_bom):
    global not_found
    # old intersect not new
    exclusive_old = difference(list(old_parent), list(new_parent))
    exclusive_old_tuples = [
        (old_bom.loc[idx, 'F/N'], old_bom.loc[idx, 'Description'])
        for idx, pn in exclusive_old
    ]
    # new intersect not old
    exclusive_new = difference(list(new_parent), list(old_parent))
    exclusive_new_tuples = [
        (new_bom.loc[idx, 'F/N'], new_bom.loc[idx, 'Description'])
        for idx, pn in exclusive_new
    ]
    for old_tuple in exclusive_old_tuples:
        match, new_tuple = is_updated(old_tuple, exclusive_new_tuples)

        old_index = exclusive_old[exclusive_old_tuples.index(old_tuple)][0]
        old_pn = old_bom.loc[old_index, "Name"]

        if match is not False:
            # Note have problem with duplicates i.e. in revg exists 5041357 where fn 1-2 and 108
            # under level 6 both are the same part number. Lucklily, revk has it duplicated as well
            # and the find number stays the same. Need to consider case where if find number is changed or
            # the second part number is removed in revk, need to identify that it is not found
            # Maybe a duplication check at the end of the update cycle on keys, if any keys are duplicated,
            # seconday one will be considered not found.

            updated_index = exclusive_new[exclusive_new_tuples.index(new_tuple)][0]
            updated_pn = new_bom.loc[updated_index, "Name"]
            old_parent[f'{old_index} {updated_pn}'] = old_parent.pop(f'{old_index} {old_pn}')
            if match == 'full':
                print(f'{old_pn, old_tuple} has been updated to {updated_pn, new_tuple}')
            else:
                print(f'Partial match for {old_pn, old_tuple} and {updated_pn, new_tuple}')
        else:
            not_found.append(f'{old_index} {old_pn}')

    for idxa, idxb in intersect(list(old_parent), list(new_parent)):
        # When getting new part number need to use new_bom since that is what everything is
        # updated to. Eg. if 5001241 becomes 5012345, 5012345 doesnt exist in the old_bom
        # so need to use idxb/ idxnew and new_bom to get correct name

        # ANother problem arises when the BOM is rearranged, where a part number i.e. 5070099
        # was moved to level 3 in revk

        # Then need to find out stuff that has been added and do the analysis in reverse
        if old_parent[f'{idxa} {new_bom.loc[idxb, "Name"]}']:
            recursive_compare(old_parent[f'{idxa} {new_bom.loc[idxb, "Name"]}'],
                              old_bom,
                              new_parent[f'{idxb} {new_bom.loc[idxb, "Name"]}'],
                              new_bom)


def bom_compare(old_parent, old_bom, new_parent, new_bom):
    global not_found

    def recursive_scan(find, obj):
        for key, value in obj.items():
            if value:
                found = recursive_scan(find, value)
                if found is not None:
                    return found
            if find in value:
                return key
        return None

    def insert(obj, key, new_key, new_value):
        for k, v in obj.items():
            if v:
                obj[k] = insert(v, key, new_key, new_value)
        if key in obj:
            obj[key][new_key] = new_value
        return obj

    def pop(obj, key):
        for key1, value in obj.items():
            if value:
                found = pop(value, key)
                if found is not None:
                    return found
        if key in obj:
            return obj.pop(key)
        return None

    def rearrange(obj, old_key, new_key):
        branch = pop(obj, old_key)
        return insert(obj, new_key, old_key, branch)

    recursive_compare(old_parent, old_bom, new_parent, new_bom)
    prev_len = 0
    while len(not_found) != prev_len:
        for item in not_found:
            pn = item.split(' ')[1]
            if pn in new_bom['Name'].values or int(item.split(' ')[1]) in new_bom['Name'].values:
                index = new_bom.loc[new_bom['Name'] == pn].index[0]
                new_key = f'{index} {pn}'

                # Found the updated part number for the nbew parent
                updated_pn = recursive_scan(new_key, new_parent).split(' ')[1]

                if updated_pn in old_bom['Name'].values:
                    # using updated part number get the new index number
                    index = old_bom.loc[old_bom['Name'] == updated_pn].index[0]
                    old_updated_key = f'{index} {updated_pn}'
                    old_parent = rearrange(old_parent, item, old_updated_key)
                    print(f'{item.split()[1]} was rearranged to be under {updated_pn}')
                else:
                    print(f'{item} was found but parent was not')

        prev_len = len(not_found)
        not_found = []
        print('recurssive compare called')
        recursive_compare(old_parent, old_bom, new_parent, new_bom)
    print(not_found)
    print(f'{len(not_found)} values are in the not_found list')


if __name__ == '__main__':
    test = [1,2,3,4,5,6,6]
    test.pop()
