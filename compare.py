"""Compare Module for BOM Comparisons

Date: 2020-7-23
Rev: A
Author: Steven Fu
Last Edit: Steven Fu
"""

from collections import Counter

import pandas as pd
from fuzzywuzzy import fuzz

from copy import deepcopy

import progressbar

class Bom:

    def __init__(self, avl_bom, parent):
        self.bom = avl_bom
        self.parent = parent
        self.parent_list = list(parent)
        self.top_pn = [Bom._split_key(pn)[1] for pn in self.parent]

    def __sub__(self, other_bom):
        """Subtrself.parent_listct between self.parent_listnother bom Object

        Only works on the highest level of parent

        Parameters:
            other_bom: Another Bom object

        Return: only the part numbers and index number (int) that exists in this self instance
        """
        return [Bom._split_key(self.parent_list[i])
                for i in range(len(self.top_pn)) if self.top_pn[i] not in other_bom.top_pn]

    @staticmethod
    def _split_key(key: str):
        """Splits the key into index and part number both in (int)"""

        return int(key.split()[0]), key.split()[1]

    def intersect(self, other_bom):
        """Find the interesection between two BOM

        Only finds the highets level intersects and returns the index and pn of current instance (self)
        """
        intersect_list = list((Counter(self.top_pn) & Counter(other_bom.top_pn)).elements())
        copy_top = self.top_pn.copy()
        copy_parent = self.parent_list.copy()
        output = []
        for item in intersect_list:
            index = copy_top.index(item)
            del copy_top[index]
            output.append(Bom._split_key(copy_parent.pop(index)))
        return sorted(output, key=lambda output: output[1])

    def immediate_parent(self, index):
        parent_level = self.bom.loc[index, 'Level'] - 1
        for idx in range(index, -1, -1):
            if self.bom.loc[idx, 'Level'] == parent_level:
                return idx, self.bom.loc[idx, 'Name']
        return None

    @staticmethod
    def zip_intersect(bom_old, bom_new):
        a = list(bom_old.parent)
        b = list(bom_new.parent)
        pa = [pn.split()[1] for pn in a]
        pb = [pn.split()[1] for pn in b]
        intersect_list = []
        for pn in range(len(pa)):
            if pa[pn] in pb:
                index_a = int(a[pn].split()[0])
                pn_a = a[pn].split()[1]

                index_b = int(b[pb.index(pa[pn])].split()[0])
                pn_b = b[pb.index(pa[pn])].split()[1]

                b.pop(pb.index(pa[pn]))
                pb.pop(pb.index(pa[pn]))

                intersect_list.append([(index_a, pn_a), (index_b, pn_b)])
        return intersect_list


class Tracker:

    COLUMNS = ['old_idx', 'old_pn', 'new_idx', 'new_pn']

    def __init__(self):
        self.full_match = pd.DataFrame(columns=self.COLUMNS)
        self.partial_match = pd.DataFrame(columns=self.COLUMNS)
        self.find_only = pd.DataFrame(columns=self.COLUMNS)
        self.not_found = set()
        self.used = []

    def append_full(self, part_old, part_new):
        self.used.append(part_new)
        append = pd.DataFrame([part_old + part_new],
                              columns=self.COLUMNS)
        self.full_match = pd.concat([self.full_match, append], ignore_index=True)

    def append_partial(self, part_old, part_new):
        self.used.append(part_new)
        append = pd.DataFrame([part_old + part_new],
                              columns=self.COLUMNS)
        self.partial_match = pd.concat([self.partial_match, append], ignore_index=True)

    def append_find_only(self, part_old, part_new):
        append = pd.DataFrame([part_old + part_new],
                              columns=self.COLUMNS)
        self.find_only = pd.concat([self.find_only, append], ignore_index=True)

    def not_found_to_df(self):
        df = pd.DataFrame(data=self.not_found, columns=['idx', 'pn'])
        return df

    def isused(self, part):
        return True if part in self.used else False

    def reset_not_found(self):
        self.not_found = set()

    def combine_found(self):
        self.full_match.insert(4, 'match_type', 'full')
        self.partial_match.insert(4, 'match_type', 'partial')
        self.find_only.insert(4, 'match_type', 'fn_only')
        combined = pd.concat([self.full_match, self.partial_match, self.find_only])
        return combined.sort_values(by=['old_idx'])


def ismatch(bom_old, part_old, bom_new, part_new, threshold_full=50, threshold_partial=80):
    old_fn, new_fn = bom_old.bom.loc[part_old[0], 'F/N'], bom_new.bom.loc[part_new[0], 'F/N']

    old_desc, new_desc = bom_old.bom.loc[part_old[0], 'Description'], bom_new.bom.loc[part_new[0], 'Description']
    match_pct = fuzz.ratio(old_desc, new_desc)

    old_type, new_type = old_desc.split('*')[0], new_desc.split('*')[0]

    if old_fn == new_fn and (match_pct >= threshold_full or old_type == new_type):
        return 'full'
    elif match_pct >= threshold_partial:
        return 'partial'
    return False


def update_part(bom_old, part_old, bom_new, exclusive_new, tracker):
    for part_new in exclusive_new:
        match_status = ismatch(bom_old, part_old, bom_new, part_new)

        if match_status == 'full' and not tracker.isused(part_new):
            print(f'{part_old[1]} {bom_old.bom.loc[part_old[0], "Description"]} updated to '
                  f'{part_new[1]} {bom_new.bom.loc[part_new[0], "Description"]}')
            tracker.append_full(part_old, part_new)

            bom_old.parent[f'{part_old[0]} {part_new[1]}'] = \
                bom_old.parent.pop(f'{part_old[0]} {part_old[1]}')
            return

    for part_new in exclusive_new:
        match_status = ismatch(bom_old, part_old, bom_new, part_new)
        if match_status == 'partial' and not tracker.isused(part_new):
            print(f'{part_old[1]} {bom_old.bom.loc[part_old[0], "Description"]} updated to '
                  f'{part_new[1]} {bom_new.bom.loc[part_new[0], "Description"]}')
            tracker.append_partial(part_old, part_new)

            bom_old.parent[f'{part_old[0]} {part_new[1]}'] = \
                bom_old.parent.pop(f'{part_old[0]} {part_old[1]}')
            return

    tracker.not_found.add(part_old)
    return


def Update(bom_old, bom_new, tracker):
    """Main updater class without rearrange

    All bom will be dealt with using the Bom class. This means the index and pn have been already split.
    In the form [index, part number]

    ## mention index numbers somewhere ##
    """
    exclusive_old = bom_old - bom_new
    exclusive_new = bom_new - bom_old

    # Checks to see if only find number has been changed
    deep_copy_parent = deepcopy(bom_new.parent_list)
    deep_copy_top = deepcopy(bom_new.top_pn)
    for part in bom_old.parent:
        key_old = Bom._split_key(part)
        if key_old not in exclusive_old:
            try:
                index = deep_copy_top.index(key_old[1])

                key_new = Bom._split_key(deep_copy_parent.pop(index))
                del deep_copy_top[index]

                if bom_old.bom.loc[key_old[0], 'F/N'] != bom_new.bom.loc[key_new[0], 'F/N']:
                    tracker.append_find_only(key_old, key_new)
            except ValueError:
                tracker.not_found.add(key_old)

    tracker.used = tracker.used + bom_new.intersect(bom_old)

    for part_old in exclusive_old:
        try:
            update_part(bom_old, part_old, bom_new, exclusive_new, tracker)
        except KeyError:
            tracker.not_found.add(part_old)

    for part_old, part_new in Bom.zip_intersect(bom_old, bom_new):
        if bom_old.parent[f'{part_old[0]} {part_new[1]}']:
            next_iter_old = Bom(bom_old.bom, bom_old.parent[f'{part_old[0]} {part_new[1]}'])
            next_iter_new = Bom(bom_new.bom, bom_new.parent[f'{part_new[0]} {part_new[1]}'])
            bom_old.parent[f'{part_old[0]} {part_new[1]}'] = Update(next_iter_old, next_iter_new, tracker)
    return bom_old.parent


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


def Rearrange(bom_old, bom_new, tracker):
    prev_len_mia = -1
    while prev_len_mia != len(tracker.not_found):

        prev_len_mia = len(tracker.not_found)
        tracker.reset_not_found()
        Update(bom_old, bom_new, tracker)

        for part in tracker.not_found:

            if part[1] in bom_new.bom['Name'].values:
                try:
                    parent_new = bom_new.immediate_parent(
                        bom_new.bom.loc[bom_new.bom['Name'] == part[1]].index[0]
                    )
                    if parent_new[1] in bom_old.bom['Name'].values:
                        parent_old_index = bom_new.bom.loc[bom_new.bom['Name'] == part[1]].index[0]
                        print(f'{part[0]} {part[1]} has been rearranged to be under {parent_old_index} {parent_new[1]}')
                        rearrange(bom_old.parent, f'{part[0]} {part[1]}', f'{parent_old_index} {parent_new[1]}')
                except TypeError:
                    continue
    for part in tracker.not_found:
        print(f'{part[1]} was not found')


if __name__ == '__main__':
    import json

    with open('Fake Bom\\treeg.json', 'r') as read:
        old_tree = json.load(read)
    old_bom = pd.read_csv('Fake Bom\\revg.csv')
    bom_old = Bom(old_bom, old_tree)

    with open('Fake Bom\\treen.json', 'r') as read:
        new_tree = json.load(read)
    new_bom = pd.read_csv('Fake Bom\\revn.csv')
    bom_new = Bom(new_bom, new_tree)

    tracker = Tracker()
    Rearrange(bom_old, bom_new, tracker)
    print(tracker.not_found_to_df())
