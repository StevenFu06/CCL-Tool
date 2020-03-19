import pandas as pd


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