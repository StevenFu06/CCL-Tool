import pandas as pd
from fuzzywuzzy import fuzz


class Bom:

    def __init__(self, avl_bom, parent_dict):
        self.bom = avl_bom
        self.parent_dict = parent_dict

    def __sub__(self, other):
        pass


class Compare:

    def __init__(self, old, new):
        self.old_bom = old
        self.new_bom = new


def compare(old, old_parent, new, new_parent):
    pass


def difference(a, b):
    """Find difference between a and b (a intersect not b)

    Effectively if a = [1,2,30,4,5] b = [24,5,1,2,3,12]
    will return [4, 30]

    :return [[idx, pn],[idx, pn],[idx, pn]...]
    """
    pa = [pn.split()[1] for pn in a]
    pb = [pn.split()[1] for pn in b]
    difference_list = []
    for pn in pa:
        if pn not in pb:
            temp_index = pa.index(pn)
            pa.pop(temp_index)
            index_a = int(a.pop(temp_index).split()[0])
            difference_list.append([index_a, pn])
    return difference_list


# def intersect(a, b):
#     pa = [pn.split()[1] for pn in a]
#     pb = [pn.split()[1] for pn in b]
#     return [
#         # Need a more robust way of switching between part numbers and idx since .index() will
#         # always return the first even if it is repeated.
#         [int(a[pa.index(pn)].split()[0]), int(b[pb.index(pn)].split()[0])]
#         for pn in pa if pn in pb
#     ]


def intersect(a, b):
    pa = [pn.split()[1] for pn in a]
    pb = [pn.split()[1] for pn in b]
    intersect_list = []
    for pn in pa:
        if pn in pb:
            temp_indexa = pa.index(pn)
            pa.pop(temp_indexa)
            temp_indexb = pb.index(pn)
            pb.pop(temp_indexb)

            index_a = int(a.pop(temp_indexa).split()[0])
            index_b = int(b.pop(temp_indexb).split()[0])
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


def recursive_compare(old_parent, old_bom, new_parent, new_bom):
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

    not_found = []
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
                print(f'{old_tuple} has been updated to {new_tuple}')
                print()
            else:
                print(f'Partial match for {old_tuple} and {new_tuple}')
                print()
        else:
            not_found.append(f'{old_index} {old_pn}')

    if not_found:
        print(f'{not_found} were not found')
        print()

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
